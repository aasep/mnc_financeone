<?php
session_start();
require_once '../../config/config.php';
require_once '../../function/function.php';
require_once '../../session_login.php';
require_once '../../session_group.php';
require_once 'Classes/PHPExcel.php';
require_once 'Classes/PHPExcel/IOFactory.php';
date_default_timezone_set("Asia/Bangkok");





$file_eksport=date('Y_m_d_H_i_s');

error_reporting(1);
logActivity("generate nop",date('Y_m_d_H_i_s'));
/*
$tanggal=$_POST['tanggal']; 
$curr_tgl=date('Y-m-d',strtotime($tanggal));

$label_txtfile=date('Ymd',strtotime($tanggal));
$tanggal_header=date('dmY',strtotime($tanggal));


$day=date('d',strtotime($tanggal));
$mon=date('M',strtotime($tanggal));
$year=date('y',strtotime($tanggal));

$mon_modal=date('n',strtotime($tanggal));
$year_modal=date('Y',strtotime($tanggal));

$label_tgl=$day."-".$mon."-".$year; // tanggal terpilih
$label_bln=$mon."-".$year; // Bulan terpilih
*/
$tahun=$_POST['tahun'];
//$kuartal=$_POST['kuartal']; // hanya bayangan
$bulan=$_POST['bulan'];
/*
switch ($kuartal) {
        case '1':
        $tanggal_awal=$tahun."-01-31";
        $tanggal=$tahun."-03-31";
        break;
        case '2':
        $tanggal_awal=$tahun."-04-30";
        $tanggal=$tahun."-06-30";
        break;
        case '3':
        $tanggal_awal=$tahun."-07-31";
        $tanggal=$tahun."-09-30";
        break;
        case '4':
        $tanggal_awal=$tahun."-10-31";
        $tanggal=$tahun."-12-31";
        break;
     
}
*/

$tmp_tgl=$tahun."-$bulan-01";
$tahun_neraca=date("Y",strtotime(date('Y-m-d',strtotime($tmp_tgl))." -1 year "));
$tgl_neraca=$tahun_neraca."-12-31";

$tgl_laba_rugi=$tahun_neraca."-$bulan-01";


//======================variable tanggal======================
$var_tgl=date("Y-m-t",strtotime(date('Y-m-d',strtotime($tmp_tgl))." 0 second "));
$before_neraca=date("Y-m-t",strtotime(date('Y-m-d',strtotime($tgl_neraca))." 0 second "));
$before_rugi_laba=date("Y-m-t",strtotime(date('Y-m-d',strtotime($tgl_laba_rugi))." 0 second "));

/*
echo $var_tgl."<br>";
echo $before_neraca."<br>";
echo $before_rugi_laba;
die();
*/
$nominal_neraca=array();
$nominal_neraca_sebelumnya=array();
$nominal_rugilaba_sebelumnya=array();

$query=" select Lap_Keu_Level_3,Lap_Keu_Level_3_Description from Referensi_Laporan_Keuangan order by Lap_Keu_Level_3 asc ";
$result=odbc_exec($connection2, $query);
        $baris=1;
        while ($row=odbc_fetch_array($result)) {
           //GET LAP_KEU_LEVEL_3
            $lap_keu_level_3=$row['Lap_Keu_Level_3'];
            
            #### CURRENT DATE ################
            $query_lf1 =" SELECT SUM(Nominal) as jml_nominal FROM DM_Journal a WITH (NOLOCK) ";
            $query_lf1.=" JOIN Referensi_GL_02 b ON b.GLNO = a.KodeGL AND b.PRODNO = a.KodeProduct ";
            $query_lf1.=" JOIN Referensi_Laporan_Keuangan c ON c.Lap_Keu_Level_3 = b.Lap_Keu_Level_3 ";
            $query_lf1.=" WHERE a.DataDate='$var_tgl' AND b.Lap_Keu_Level_3='$lap_keu_level_3' ";
            $result_lf1=odbc_exec($connection2, $query_lf1);
            $row1=odbc_fetch_array($result_lf1);
            $nominal_neraca["$baris"]=$row1['jml_nominal'];

            
            //array_push($nominal_neraca, $row1['jml_nominal']);
            //echo $query_lf1;
            //die();
            #### BEFORE DATE  NERACA ################
            $query_lf2 =" SELECT SUM(Nominal) as jml_nominal FROM DM_Journal a WITH (NOLOCK) ";
            $query_lf2.=" JOIN Referensi_GL_02 b ON b.GLNO = a.KodeGL AND b.PRODNO = a.KodeProduct ";
            $query_lf2.=" JOIN Referensi_Laporan_Keuangan c ON c.Lap_Keu_Level_3 = b.Lap_Keu_Level_3 ";
            $query_lf2.=" WHERE a.DataDate='$before_neraca' AND b.Lap_Keu_Level_3='$lap_keu_level_3' ";
            $result_lf2=odbc_exec($connection2, $query_lf2);
            $row2=odbc_fetch_array($result_lf2);
            $nominal_neraca_sebelumnya["$baris"]=$row2['jml_nominal'];
            //array_push($nominal_neraca_sebelumnya, $row2['jml_nominal']);
            #### BEFORE DATE  RUGI LABA  ################
            $query_lf3 =" SELECT SUM(Nominal) as jml_nominal FROM DM_Journal a WITH (NOLOCK) ";
            $query_lf3.=" JOIN Referensi_GL_02 b ON b.GLNO = a.KodeGL AND b.PRODNO = a.KodeProduct ";
            $query_lf3.=" JOIN Referensi_Laporan_Keuangan c ON c.Lap_Keu_Level_3 = b.Lap_Keu_Level_3 ";
            $query_lf3.=" WHERE a.DataDate='$before_rugi_laba' AND b.Lap_Keu_Level_3='$lap_keu_level_3' ";
            $result_lf3=odbc_exec($connection2, $query_lf3);
            $row3=odbc_fetch_array($result_lf3);
            $nominal_rugilaba_sebelumnya["$baris"]=$row3['jml_nominal'];
            //array_push($nominal_rugilaba_sebelumnya, $row3['jml_nominal']);


            $baris++;

        }





//var_dump($nominal_neraca);
//echo "<br>";
//var_dump($nominal_neraca_sebelumnya);
//echo "<br>";
//var_dump($nominal_rugilaba_sebelumnya);

//die();

// Create new PHPExcel object
$objPHPExcel = new PHPExcel();

$styleArrayFontBold = array('font' => array('bold'  => true,'color' => array('rgb' => ''),'size'  => 11,'name'  => 'Calibri'));
$styleArrayFontBold2 = array('font' => array('bold'  => true,'color' => array('rgb' => ''),'size'  => 13,'name'  => 'Calibri'));
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

$objPHPExcel->getActiveSheet()->getStyle('A1:D3')->applyFromArray($styleArrayFontBold2);
$objPHPExcel->getActiveSheet()->getStyle('A26:D26')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('A7:D7')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('A40:D42')->applyFromArray($styleArrayFontBold2);


$objPHPExcel->getActiveSheet()->getStyle('A85:A87')->applyFromArray($styleArrayFontBold2);
$objPHPExcel->getActiveSheet()->getStyle('B91')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('B119')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('B120')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('B124')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('B124:D128')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('B133')->applyFromArray($styleArrayFontBold);
//$objPHPExcel->getActiveSheet()->getStyle('B133')->applyFromArray($styleArrayFontBold);


$objPHPExcel->getActiveSheet()->getStyle('A45:D47')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('A63:D64')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('A74:D75')->applyFromArray($styleArrayFontBold);





$objPHPExcel->getActiveSheet()->getStyle('A4:D4')->applyFromArray($styleArrayAlignment1);

//=======BORDER
$styleArrayBorder1 = array('borders' => array('allborders' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('argb' => ''),),),);
$styleArrayBorder2 = array('borders' => array('outline' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('argb' => ''),),),);
$styleArrayBorder3 = array('borders' => array('bottom' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('argb' => ''),),),);

$objPHPExcel->getActiveSheet()->getStyle('B7:D26')->applyFromArray($styleArrayBorder1);
$objPHPExcel->getActiveSheet()->getStyle('B46:D75')->applyFromArray($styleArrayBorder1);
$objPHPExcel->getActiveSheet()->getStyle('B91:D134')->applyFromArray($styleArrayBorder1);
//FILL COLOR
//$objPHPExcel->getActiveSheet()->getStyle('A10:E10')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('D3D3D3');

//DIMENSION D

$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(5);
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(70);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(25);
$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(25);


// Create a first sheet, representing sales data

$objPHPExcel->setActiveSheetIndex(0);
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('A1:D1');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('A2:D2');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('A3:D3');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('A4:D4');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('A28:D28');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('A29:D29');

$objPHPExcel->setActiveSheetIndex(0)->mergeCells('A40:D40');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('A41:D41');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('A42:D42');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('A43:D43');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('A78:D78');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('A79:D79');

$objPHPExcel->setActiveSheetIndex(0)->mergeCells('A136:D136');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('A137:D137');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('A85:D85');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('A86:D86');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('A87:D87');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('A88:D88');


$objPHPExcel->setActiveSheetIndex(0)->mergeCells('A78:D78');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('A79:D79');

$objPHPExcel->getActiveSheet()->getStyle('A1:D4')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A1:D4')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A9:D26')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('C7:D7')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A28:A29')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A45:D75')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

$objPHPExcel->getActiveSheet()->getStyle('A40:D43')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A78:D79')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);


$objPHPExcel->getActiveSheet()->getStyle('A85:D88')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A136:D137')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

$objPHPExcel->getActiveSheet()->getStyle('A1:Z200')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('E1:Z200')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');
//$objPHPExcel->getActiveSheet()->getStyle('A27:Z100')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('B7:D7')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('808080');


$objPHPExcel->getActiveSheet()->setCellValue('A1', 'PT BANK MNC INTERNASIONAL TBK.');
$objPHPExcel->getActiveSheet()->setCellValue('A2', "LAPORAN POSISI KEUANGAN");
$objPHPExcel->getActiveSheet()->setCellValue('A3', "Per $tanggal_awal ");
$objPHPExcel->getActiveSheet()->setCellValue('A4', 'Disajikan Dalam Jutaan Rupiah Kecuali Dinyatakan yang lain');
$objPHPExcel->getActiveSheet()->setCellValue('B7', 'ASET');
$objPHPExcel->getActiveSheet()->setCellValue('C7', "$var_tgl");
$objPHPExcel->getActiveSheet()->setCellValue('D7', $before_neraca);

for ($i=9; $i <=29 ; $i++) { 
$objPHPExcel->getActiveSheet(0)->getRowDimension($i)->setRowHeight(20);
$objPHPExcel->getActiveSheet()->getStyle("C$i:D$i")->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
}
for ($i=45; $i <=75 ; $i++) { 
$objPHPExcel->getActiveSheet(0)->getRowDimension($i)->setRowHeight(20);
$objPHPExcel->getActiveSheet()->getStyle("C$i:D$i")->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
}
for ($i=91; $i <=134 ; $i++) { 
$objPHPExcel->getActiveSheet(0)->getRowDimension($i)->setRowHeight(18);
$objPHPExcel->getActiveSheet()->getStyle("C$i:D$i")->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
}

$objPHPExcel->getActiveSheet()->setCellValue("B9", "Kas");
$objPHPExcel->getActiveSheet()->setCellValue("B10", "Giro Pada Bank Indonesia");
$objPHPExcel->getActiveSheet()->setCellValue("B11", "Giro Pada Bank Lain - Pihak Ketiga");
$objPHPExcel->getActiveSheet()->setCellValue("B12", "Penempatan Pada Bank Indonesia dan Bank Lain - Pihak ketiga");
$objPHPExcel->getActiveSheet()->setCellValue("B13", "Efek-Efek Pihak ketiga");
$objPHPExcel->getActiveSheet()->setCellValue("B14", "Tagihan Derivatif - Pihak ketiga");
$objPHPExcel->getActiveSheet()->setCellValue("B15", "Kredit Yang Diberikan");
$objPHPExcel->getActiveSheet()->setCellValue("B16", " - Pihak Berelasi");
$objPHPExcel->getActiveSheet()->setCellValue("B17", " - Pihak Ketiga");
$objPHPExcel->getActiveSheet()->setCellValue("B18", " - Cadangan Kerugian Penurunan Nilai");
$objPHPExcel->getActiveSheet()->setCellValue("B19", " - Jumlah");
$objPHPExcel->getActiveSheet()->setCellValue("B20", "Tagihan Akseptasi - Pihak ketiga");
$objPHPExcel->getActiveSheet()->setCellValue("B21", "Biaya dibayar dimuka");
$objPHPExcel->getActiveSheet()->setCellValue("B22", "Aset Tetap - Bersih ");
$objPHPExcel->getActiveSheet()->setCellValue("B23", "Aset Pajak Tangguhan - bersih");
$objPHPExcel->getActiveSheet()->setCellValue("B24", "Aset Tetap tidak berwujud - Bersih");
$objPHPExcel->getActiveSheet()->setCellValue("B25", "Aset Lain-lain - bersih");
$objPHPExcel->getActiveSheet()->setCellValue("B26", "JUMLAH ASET");
//$objPHPExcel->getActiveSheet()->setCellValue("B27", "JUMLAH ASET---------");
//$objPHPExcel->getActiveSheet()->setCellValue("B28", "JUMLAH ASET---------");
$objPHPExcel->getActiveSheet()->setCellValue("A28", "Lihat catatan atas laporan keuangan yang merupakan bagian tak terpisahkan ");
$objPHPExcel->getActiveSheet()->setCellValue("A29", "dari laporan keuangan secara keseluruhan ");


$objPHPExcel->getActiveSheet()->setCellValue('A40', 'PT BANK MNC INTERNASIONAL TBK.');
$objPHPExcel->getActiveSheet()->setCellValue('A41', "LAPORAN POSISI KEUANGAN");
$objPHPExcel->getActiveSheet()->setCellValue('A42', "Per $tanggal_awal ");
$objPHPExcel->getActiveSheet()->setCellValue('A43', 'Disajikan Dalam Jutaan Rupiah Kecuali Dinyatakan yang lain');

$objPHPExcel->getActiveSheet()->setCellValue('C45', "$var_tgl");
$objPHPExcel->getActiveSheet()->setCellValue('D45', $before_neraca);
     

$objPHPExcel->getActiveSheet()->setCellValue('B46', 'LIABILITAS DAN EKUITAS');
$objPHPExcel->getActiveSheet()->setCellValue('B47', 'LIABILITAS');

$objPHPExcel->getActiveSheet()->setCellValue('B48', 'Liabilitas Segera');
$objPHPExcel->getActiveSheet()->setCellValue('B49', 'Simpanan');
$objPHPExcel->getActiveSheet()->setCellValue('B50', '  - Pihak Berelasi');
$objPHPExcel->getActiveSheet()->setCellValue('B51', '  - Pihak Ketiga');
$objPHPExcel->getActiveSheet()->setCellValue('B52', 'Jumlah');
$objPHPExcel->getActiveSheet()->setCellValue('B53', 'Simpanan dari Bank Lain - Pihak Ketiga');
$objPHPExcel->getActiveSheet()->setCellValue('B54', 'Liabilitas Derivatif - Pihak Ketiga');
$objPHPExcel->getActiveSheet()->setCellValue('B55', 'Liabilitas atas efek efek yang dijual dengan janji dijual kembali');
$objPHPExcel->getActiveSheet()->setCellValue('B56', 'Liabilitas Akseptasi - Pihak Ketiga');
$objPHPExcel->getActiveSheet()->setCellValue('B57', 'Obligasi Konversi');
$objPHPExcel->getActiveSheet()->setCellValue('B58', 'Pinjaman Diterima - Pihak Ketiga ');
$objPHPExcel->getActiveSheet()->setCellValue('B59', 'Utang Pajak');
$objPHPExcel->getActiveSheet()->setCellValue('B60', 'Liabilitas Imbalan pasca kerja');
$objPHPExcel->getActiveSheet()->setCellValue('B61', 'Beban yang masih harus dibayar');
$objPHPExcel->getActiveSheet()->setCellValue('B62', 'Liabilitas lain-lain');


$objPHPExcel->getActiveSheet()->setCellValue('B63', 'JUMLAH KEWAJIBAN');
$objPHPExcel->getActiveSheet()->setCellValue('B64', 'EKUITAS');
$objPHPExcel->getActiveSheet()->setCellValue('B65', 'Modal saham dengan nilai nominal Rp 100 per saham');
$objPHPExcel->getActiveSheet()->setCellValue('B66', ' - Modal dasar - 60.000.000 saham');
$objPHPExcel->getActiveSheet()->setCellValue('B67', ' - Modal ditempatkan dan disetor 1.503.232.706,- bersih');
$objPHPExcel->getActiveSheet()->setCellValue('B68', 'Tambahan modal disetor - bersih');
$objPHPExcel->getActiveSheet()->setCellValue('B69', 'Komponen Ekuitas Lainnya - Perubahan ');
$objPHPExcel->getActiveSheet()->setCellValue('B70', ' - nilai wajar efek tersedia untuk dijual - bersih');
$objPHPExcel->getActiveSheet()->setCellValue('B71', 'Saldo Laba (Rugi) ');
$objPHPExcel->getActiveSheet()->setCellValue('B72', ' - Telah ditentukan penggunaannya');
$objPHPExcel->getActiveSheet()->setCellValue('B73', ' - Belum ditentukan penggunaannya');
$objPHPExcel->getActiveSheet()->setCellValue('B74', 'JUMLAH EKUITAS');
$objPHPExcel->getActiveSheet()->setCellValue('B75', 'JUMLAH KEWAJIBAN DAN EKUITAS');

$objPHPExcel->getActiveSheet()->setCellValue("A78", "Lihat catatan atas laporan keuangan yang merupakan bagian tak terpisahkan ");
$objPHPExcel->getActiveSheet()->setCellValue("A79", "dari laporan keuangan secara keseluruhan ");

$objPHPExcel->getActiveSheet()->setCellValue('A85', 'PT BANK MNC INTERNASIONAL TBK.');
$objPHPExcel->getActiveSheet()->setCellValue('A86', "LAPORAN LABA RUGI");
$objPHPExcel->getActiveSheet()->setCellValue('A87', "Untuk Periode yang berakhir pada tanggal  $tanggal_awal ");
$objPHPExcel->getActiveSheet()->setCellValue('A88', 'Disajikan Dalam Jutaan Rupiah Kecuali Dinyatakan yang lain');

$objPHPExcel->getActiveSheet()->setCellValue('C90', "$var_tgl");
$objPHPExcel->getActiveSheet()->setCellValue('D90', $before_rugi_laba);

$objPHPExcel->getActiveSheet()->setCellValue('B91', 'PENDAPATAN DAN BEBAN OPERASIONAL');
$objPHPExcel->getActiveSheet()->setCellValue('B92', 'Pendapatan Bunga');
$objPHPExcel->getActiveSheet()->setCellValue('B93', '  Bunga yang diperoleh');
$objPHPExcel->getActiveSheet()->setCellValue('B94', '  Komisi dan Fee dari kredit yang diberikan');
$objPHPExcel->getActiveSheet()->setCellValue('B95', '  Jumlah Pendapatan Bunga');

$objPHPExcel->getActiveSheet()->setCellValue('B96', 'Beban Bunga');
$objPHPExcel->getActiveSheet()->setCellValue('B97', '  Bunga');
$objPHPExcel->getActiveSheet()->setCellValue('B98', '  Provisi dan komisi yang harus dibayar');
$objPHPExcel->getActiveSheet()->setCellValue('B99', '  Jumlah Bebean Bunga');

$objPHPExcel->getActiveSheet()->setCellValue('B100', 'Pendapatan Bunga - Bersih');
$objPHPExcel->getActiveSheet()->setCellValue('B101', 'Pendapatan Operasional Lainnya');
$objPHPExcel->getActiveSheet()->setCellValue('B102', '  Pendapatan transaksi valuta asing-bersih');
$objPHPExcel->getActiveSheet()->setCellValue('B103', '  Keuntungan bersih penjualan efek');
$objPHPExcel->getActiveSheet()->setCellValue('B104', '  Provisi komisi dan fee selain kredit - bersih');
$objPHPExcel->getActiveSheet()->setCellValue('B105', '  Penerimaan kembali kredit yang dihapus buku ');
$objPHPExcel->getActiveSheet()->setCellValue('B106', '  Lain - lain');
$objPHPExcel->getActiveSheet()->setCellValue('B107', '  Jumlah Pendapatan Operasi Lainnya');


$objPHPExcel->getActiveSheet()->setCellValue('B108', 'Beban (Pemulihan) kerugian penurunan nilai');
$objPHPExcel->getActiveSheet()->setCellValue('B109', '  Aset Keuangan');
$objPHPExcel->getActiveSheet()->setCellValue('B110', '  Aset Non Keuangan');
$objPHPExcel->getActiveSheet()->setCellValue('B111', '  Jumlah Beban Kerugian Penurunan Nilai');
$objPHPExcel->getActiveSheet()->setCellValue('B112', 'Beban Operasional lainnya');
$objPHPExcel->getActiveSheet()->setCellValue('B113', '  Umum dan administrasi ');
$objPHPExcel->getActiveSheet()->setCellValue('B114', '  Tenaga Kerja');
$objPHPExcel->getActiveSheet()->setCellValue('B115', '  Beban Pensiun dan imbalan pasca kerja lainnya');
$objPHPExcel->getActiveSheet()->setCellValue('B116', '  Lainnya');
$objPHPExcel->getActiveSheet()->setCellValue('B117', '  Jumlah Beban Operasional lainnya');


$objPHPExcel->getActiveSheet()->setCellValue('B118', 'Beban Operasional lainnya - Bersih');
$objPHPExcel->getActiveSheet()->setCellValue('B119', 'LABA (RUGI) OPERASIONAL');
$objPHPExcel->getActiveSheet()->setCellValue('B120', 'PENDAPATAN (BEBAN) NON OPERASIONAL');
$objPHPExcel->getActiveSheet()->setCellValue('B121', '  Hasil Sewa');
$objPHPExcel->getActiveSheet()->setCellValue('B122', '  Keuntungan penjualan dan pengahapusan aset tetap');
$objPHPExcel->getActiveSheet()->setCellValue('B123', '  Lainnya - bersih');
$objPHPExcel->getActiveSheet()->setCellValue('B124', 'PENDAPATAN (BEBAN) NON OPERASIONAL - BERSIH');
$objPHPExcel->getActiveSheet()->setCellValue('B125', 'LABA (RUGI) SEBELUM PAJAK PENGHASILAN');
$objPHPExcel->getActiveSheet()->setCellValue('B126', 'MANFAAT (BEBAN) PAJAK');
$objPHPExcel->getActiveSheet()->setCellValue('B127', 'LABA (RUGI) BERSIH TAHUN BERJALAN');

$objPHPExcel->getActiveSheet()->setCellValue('B128', 'PENDAPATAN KOMPREHENSIF LAIN');
$objPHPExcel->getActiveSheet()->setCellValue('B129', 'Perubahan nilai wajar efek tersedia untuk dijual');
$objPHPExcel->getActiveSheet()->setCellValue('B130', 'manfaat (Beban) pajak terkait dengan komponen');
$objPHPExcel->getActiveSheet()->setCellValue('B131', '   Pendapatan Komprehensif Lainnya');
$objPHPExcel->getActiveSheet()->setCellValue('B132', 'Jumlah pendapatan komprehensif tahun berjalan setelah pajak');
$objPHPExcel->getActiveSheet()->setCellValue('B133', 'TOTAL LABA (RUGI) KOMPREHENSIF TAHUN BERJALAN');
$objPHPExcel->getActiveSheet()->setCellValue('B134', 'Laba (Rugi) per Saham');

$objPHPExcel->getActiveSheet()->setCellValue("A136", "Lihat catatan atas laporan keuangan yang merupakan bagian tak terpisahkan ");
$objPHPExcel->getActiveSheet()->setCellValue("A137", "dari laporan keuangan secara keseluruhan ");
    
$baris=9;
for ($i=1; $i <=6 ; $i++) { 
$objPHPExcel->getActiveSheet()->setCellValue("C$baris", floatval($nominal_neraca["$i"]));
$baris++;
}

$objPHPExcel->getActiveSheet()->setCellValue("C17", floatval($nominal_neraca["7"]));
$objPHPExcel->getActiveSheet()->setCellValue("C18", floatval($nominal_neraca["8"]));
$objPHPExcel->getActiveSheet()->setCellValue("C19", "=SUM(C16:C18)");

$baris2=20;
for ($i=9; $i <=12 ; $i++) { 
$objPHPExcel->getActiveSheet()->setCellValue("C$baris2", floatval($nominal_neraca["$i"]));
$baris2++;
}

$objPHPExcel->getActiveSheet()->setCellValue("C24", floatval($nominal_neraca["14"]));
$objPHPExcel->getActiveSheet()->setCellValue("C25", floatval($nominal_neraca["13"]));

$objPHPExcel->getActiveSheet()->setCellValue("C26", "=(SUM(C9:C14)+C19+SUM(C20:C25))");

//============ neraca tahun sebelumnya

$baris=9;
for ($i=1; $i <=6 ; $i++) { 
$objPHPExcel->getActiveSheet()->setCellValue("D$baris", floatval($nominal_neraca_sebelumnya["$i"]));
$baris++;
}

$objPHPExcel->getActiveSheet()->setCellValue("D17", floatval($nominal_neraca_sebelumnya["7"]));
$objPHPExcel->getActiveSheet()->setCellValue("D18", floatval($nominal_neraca_sebelumnya["8"]));
$objPHPExcel->getActiveSheet()->setCellValue("D19", "=SUM(D16:D18)");

$baris2=20;
for ($i=9; $i <=12 ; $i++) { 
$objPHPExcel->getActiveSheet()->setCellValue("D$baris2", floatval($nominal_neraca_sebelumnya["$i"]));
$baris2++;
}

$objPHPExcel->getActiveSheet()->setCellValue("D24", floatval($nominal_neraca_sebelumnya["14"]));
$objPHPExcel->getActiveSheet()->setCellValue("D25", floatval($nominal_neraca_sebelumnya["13"]));

$objPHPExcel->getActiveSheet()->setCellValue("D26", "=(SUM(D9:D14)+D19+SUM(D20:D25))");


//======================  Neraca bagian  2 ===========================

$objPHPExcel->getActiveSheet()->setCellValue("C48", floatval($nominal_neraca["15"]));
$objPHPExcel->getActiveSheet()->setCellValue("D48", floatval($nominal_neraca_sebelumnya["15"]));

$objPHPExcel->getActiveSheet()->setCellValue("C51", floatval($nominal_neraca["16"]));
$objPHPExcel->getActiveSheet()->setCellValue("D51", floatval($nominal_neraca_sebelumnya["16"]));
$baris=53;
for ($i=17; $i <=28 ; $i++) { 
if ($i=='21' || $i=='25') {
} else {   
$objPHPExcel->getActiveSheet()->setCellValue("C$baris", floatval($nominal_neraca["$i"]));
$objPHPExcel->getActiveSheet()->setCellValue("D$baris", floatval($nominal_neraca_sebelumnya["$i"]));
$baris++;
}
}

$objPHPExcel->getActiveSheet()->setCellValue("C63", "=C48+C52+SUM(C53:C62)");
$objPHPExcel->getActiveSheet()->setCellValue("D63", "=D48+D52+SUM(D53:D62)");


$baris=66;
for ($i=29; $i <=32 ; $i++) { 
$objPHPExcel->getActiveSheet()->setCellValue("C$baris", floatval($nominal_neraca["$i"]));
$objPHPExcel->getActiveSheet()->setCellValue("D$baris", floatval($nominal_neraca_sebelumnya["$i"]));
$baris++;
}
$objPHPExcel->getActiveSheet()->setCellValue("C71", floatval($nominal_neraca["33"]));
$objPHPExcel->getActiveSheet()->setCellValue("D71", floatval($nominal_neraca_sebelumnya["33"]));


$objPHPExcel->getActiveSheet()->setCellValue("C74", "=SUM(C65:C73)");
$objPHPExcel->getActiveSheet()->setCellValue("D74", "=SUM(D65:D73)");


$objPHPExcel->getActiveSheet()->setCellValue("C75", "=C63+C74");
$objPHPExcel->getActiveSheet()->setCellValue("D75", "=D63+D74");




############# //SHEET 1 RUGI - LABA ############################################
$index=34;
for ($i=93; $i <=123 ; $i++) { 
if ($i=='95' || $i=='96' || $i=='99' || $i=='100' || $i=='101' || $i=='107' || $i=='108' || $i=='111' || $i=='112' || $i=='117' || $i=='118' || $i=='119' || $i=='120') {
} else {   
$objPHPExcel->getActiveSheet()->setCellValue("C$i", floatval($nominal_neraca["$index"]));
$objPHPExcel->getActiveSheet()->setCellValue("D$i", floatval($nominal_rugilaba_sebelumnya["$index"]));
$index++;
}
}

$objPHPExcel->getActiveSheet()->setCellValue("C95", "=SUM(C93:C94)");
$objPHPExcel->getActiveSheet()->setCellValue("D95", "=SUM(D93:D94)");

$objPHPExcel->getActiveSheet()->setCellValue("C99", "=SUM(C97:C98)");
$objPHPExcel->getActiveSheet()->setCellValue("D99", "=SUM(D97:D98)");

$objPHPExcel->getActiveSheet()->setCellValue("C100", "=C95+C99");
$objPHPExcel->getActiveSheet()->setCellValue("D100", "=D95+D99");

$objPHPExcel->getActiveSheet()->setCellValue("C107", "=SUM(C102:C106)");
$objPHPExcel->getActiveSheet()->setCellValue("D107", "=SUM(D102:D106)");

$objPHPExcel->getActiveSheet()->setCellValue("C111", "=SUM(C109:C110)");
$objPHPExcel->getActiveSheet()->setCellValue("D111", "=SUM(D109:D110)");

$objPHPExcel->getActiveSheet()->setCellValue("C117", "=SUM(C113:C116)");
$objPHPExcel->getActiveSheet()->setCellValue("D117", "=SUM(D113:D116)");
$objPHPExcel->getActiveSheet()->setCellValue("C118", "=+C111+C117");
$objPHPExcel->getActiveSheet()->setCellValue("D118", "=+D111+D117");
$objPHPExcel->getActiveSheet()->setCellValue("C119", "=+C100+C107+C118");
$objPHPExcel->getActiveSheet()->setCellValue("D119", "=+D100+D107+D118");

$objPHPExcel->getActiveSheet()->setCellValue("C124", "=SUM(C121:C123)");
$objPHPExcel->getActiveSheet()->setCellValue("D124", "=SUM(D121:D123)");
$objPHPExcel->getActiveSheet()->setCellValue("C125", "=C119+C124");
$objPHPExcel->getActiveSheet()->setCellValue("D125", "=D119+D124");

$objPHPExcel->getActiveSheet()->setCellValue("C127", "=C125+C126");
$objPHPExcel->getActiveSheet()->setCellValue("D127", "=D125+D126");

$objPHPExcel->getActiveSheet()->setCellValue("C132", "=SUM(C130:C131)");
$objPHPExcel->getActiveSheet()->setCellValue("D132", "=SUM(D130:D131)");
$objPHPExcel->getActiveSheet()->setCellValue("C133", "=+C132");
$objPHPExcel->getActiveSheet()->setCellValue("D133", "=+D132");




$objPHPExcel->getActiveSheet()->setTitle('Neraca & Laba Rugi');

// SHEET KE 2 ======================================================================================
$objPHPExcel->createSheet();
$objPHPExcel->setActiveSheetIndex(1); 

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

//$objPHPExcel->getActiveSheet()->getStyle('C7:G15')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');

$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(5);
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(130);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(25);


$objPHPExcel->setActiveSheetIndex(1)->mergeCells('A1:C1');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('A2:C2');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('A3:C3');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('A4:C4');

$objPHPExcel->setActiveSheetIndex(1)->mergeCells('A38:C38');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('A39:C39');

$objPHPExcel->setActiveSheetIndex(1)->mergeCells('A50:C50');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('A51:C51');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('A52:C52');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('A53:C53');

$objPHPExcel->setActiveSheetIndex(1)->mergeCells('A86:C86');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('A87:C87');

$objPHPExcel->setActiveSheetIndex(1)->mergeCells('A100:C100');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('A101:C101');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('A102:C102');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('A103:C103');
//BOLD
$objPHPExcel->getActiveSheet()->getStyle('A1:C3')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('B8')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('A50:A52')->applyFromArray($styleArrayFontBold);

$objPHPExcel->getActiveSheet()->getStyle('B54:C55')->applyFromArray($styleArrayFontBold);

$objPHPExcel->getActiveSheet()->getStyle('B72:C73')->applyFromArray($styleArrayFontBold);

$objPHPExcel->getActiveSheet()->getStyle('B81:B82')->applyFromArray($styleArrayFontBold);

$objPHPExcel->getActiveSheet()->getStyle('A100:A102')->applyFromArray($styleArrayFontBold);

$objPHPExcel->getActiveSheet()->getStyle('B106')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('B110')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('B120:B121')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('B131')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('B141')->applyFromArray($styleArrayFontBold);

//BORDER=============
$objPHPExcel->getActiveSheet()->getStyle('B8:C30')->applyFromArray($styleArrayBorder1);
$objPHPExcel->getActiveSheet()->getStyle('B54:C82')->applyFromArray($styleArrayBorder1);
$objPHPExcel->getActiveSheet()->getStyle('B106:C141')->applyFromArray($styleArrayBorder1);


$objPHPExcel->getActiveSheet()->getStyle('A1:C4')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A38:C39')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A50:A53')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A86:A87')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A100:A103')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

$objPHPExcel->getActiveSheet()->getStyle('B25:B27')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

$objPHPExcel->getActiveSheet()->getStyle('A53')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

$objPHPExcel->getActiveSheet()->getStyle('A1:Z200')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');

$objPHPExcel->getActiveSheet(1)->getRowDimension(25)->setRowHeight(30);
$objPHPExcel->getActiveSheet(1)->getRowDimension(26)->setRowHeight(30);
$objPHPExcel->getActiveSheet(1)->getRowDimension(27)->setRowHeight(30);


$objPHPExcel->getActiveSheet()->setCellValue('A1', 'PT BANK MNC INTERNASIONAL TBK.');
$objPHPExcel->getActiveSheet()->setCellValue('A2', "LAPORAN POSISI KEUANGAN");
$objPHPExcel->getActiveSheet()->setCellValue('A3', "Per $tanggal_awal (Tidak diaudit)  ");
$objPHPExcel->getActiveSheet()->setCellValue('A4', 'Disajikan Dalam Jutaan Rupiah Kecuali Dinyatakan yang lain');



$objPHPExcel->getActiveSheet()->setCellValue('B8', 'ASET');
$objPHPExcel->getActiveSheet()->setCellValue('C8', " $tanggal_awal ");

$objPHPExcel->getActiveSheet()->setCellValue('B10', 'Kas');
$objPHPExcel->getActiveSheet()->setCellValue('B11', 'Giro Pada Bank Indonesia');
$objPHPExcel->getActiveSheet()->setCellValue('B12', 'Giro Pada Bank Lain ');
$objPHPExcel->getActiveSheet()->setCellValue('B13', 'Penempatan Pada Bank Indonesia dan Bank Lain');
$objPHPExcel->getActiveSheet()->setCellValue('B14', 'Efek-Efek Pihak ketiga');
$objPHPExcel->getActiveSheet()->setCellValue('B15', 'Cadangan Kerugian Penurunan Nilai');
$objPHPExcel->getActiveSheet()->setCellValue('B16', 'Total Efek Efek');
$objPHPExcel->getActiveSheet()->setCellValue('B17', 'Tagihan Derivatif ');
$objPHPExcel->getActiveSheet()->setCellValue('B18', 'Kredit Yang Diberikan ');
$objPHPExcel->getActiveSheet()->setCellValue('B19', 'Pihak Berelasi');
$objPHPExcel->getActiveSheet()->setCellValue('B20', 'Pihak Ketiga');
$objPHPExcel->getActiveSheet()->setCellValue('B21', 'Jumlah Kredit Yang Diberikan');
$objPHPExcel->getActiveSheet()->setCellValue('B22', 'Cadangan Kerugian Penurunan Nilai');
$objPHPExcel->getActiveSheet()->setCellValue('B23', 'Tagihan Akseptasi ');
$objPHPExcel->getActiveSheet()->setCellValue('B24', 'Pendapatan bunga yang masih akan diterima');
$objPHPExcel->getActiveSheet()->setCellValue('B25', 'Aset Tetap - setelah dikurangi akumulasi penyusutan sebesar Rp 86,253,909,274  pada 31 Oktober  2015 ');
$objPHPExcel->getActiveSheet()->setCellValue('B26', 'Aset Tidak Berwujud  - setelah dikurangi akumulasi penyusutan sebesar Rp. 64,267,865,294 pada  31 Oktober 2015');
$objPHPExcel->getActiveSheet()->setCellValue('B27', 'Agunan yang diambil alih - setelah dikurangi Cadangan Kerugian Penurunan Nilai sebesar Rp 19,401,000,292 pada 31 Oktober 2015. ');

$objPHPExcel->getActiveSheet()->setCellValue('B28', 'Beban dibayar dimuka');
$objPHPExcel->getActiveSheet()->setCellValue('B29', 'Aset Lain-lain - bersih');
$objPHPExcel->getActiveSheet()->setCellValue('B30', 'JUMLAH ASET ');

$objPHPExcel->getActiveSheet()->setCellValue('A38', 'Lihat catatan atas laporan keuangan yang merupakan bagian tak terpisahkan ');
$objPHPExcel->getActiveSheet()->setCellValue('A39', 'dari laporan keuangan secara keseluruhan ');

//=========================================================================================================================================================================


$objPHPExcel->getActiveSheet()->setCellValue('A50', 'PT BANK MNC INTERNASIONAL TBK.');
$objPHPExcel->getActiveSheet()->setCellValue('A51', "LAPORAN POSISI KEUANGAN");
$objPHPExcel->getActiveSheet()->setCellValue('A52', "Per $tanggal_awal (Tidak diaudit)  ");
$objPHPExcel->getActiveSheet()->setCellValue('A53', 'Disajikan Dalam Jutaan Rupiah Kecuali Dinyatakan yang lain');
$objPHPExcel->getActiveSheet(0)->getRowDimension(53)->setRowHeight(50);


$objPHPExcel->getActiveSheet()->setCellValue('B54', 'LIABILITAS DAN EKUITAS');
$objPHPExcel->getActiveSheet()->setCellValue('C54', $tanggal_awal);

$objPHPExcel->getActiveSheet()->setCellValue('B55', 'LIABILITAS');
$objPHPExcel->getActiveSheet()->setCellValue('B56', 'Liabilitas Segera ');
$objPHPExcel->getActiveSheet()->setCellValue('B57', 'Simpanan');
$objPHPExcel->getActiveSheet()->setCellValue('B58', 'Pihak Berelasi');
$objPHPExcel->getActiveSheet()->setCellValue('B59', 'Pihak Ketiga');
$objPHPExcel->getActiveSheet()->setCellValue('B60', 'Jumlah Simpanan');
$objPHPExcel->getActiveSheet()->setCellValue('B61', 'Simpanan dari Bank lain');

$objPHPExcel->getActiveSheet()->setCellValue('B62', 'Surat Berharga yang dijual dengan janji dibeli kembali');
$objPHPExcel->getActiveSheet()->setCellValue('B63', 'Liabilitas Derivatif');
$objPHPExcel->getActiveSheet()->setCellValue('B64', 'Liabilitas Akseptasi');
$objPHPExcel->getActiveSheet()->setCellValue('B65', 'Pinjaman yang Diterima');
$objPHPExcel->getActiveSheet()->setCellValue('B66', 'Estimasi kerugian komitmen dan kontijensi');
$objPHPExcel->getActiveSheet()->setCellValue('B67', 'Hutang Pajak');
$objPHPExcel->getActiveSheet()->setCellValue('B68', 'Komponen liabilitas dari Obligasi Wajib Konversi');
$objPHPExcel->getActiveSheet()->setCellValue('B69', 'Bunga masih harus dibayar');
$objPHPExcel->getActiveSheet()->setCellValue('B70', 'Liabilitas Imbalan pasca kerja');
$objPHPExcel->getActiveSheet()->setCellValue('B71', 'Liabilitas lain-lain');
$objPHPExcel->getActiveSheet()->setCellValue('B72', 'JUMLAH LIABILITAS  ');
$objPHPExcel->getActiveSheet()->setCellValue('B73', 'EKUITAS ');
$objPHPExcel->getActiveSheet()->setCellValue('B74', 'Modal saham dengan nilai nominal Rp 100 per saham (nilai penuh) Modal dasar - 20.000.000.000 lembar saham Modal ditempatkan dan disetor - penuh');
$objPHPExcel->getActiveSheet()->setCellValue('B75', 'Tambahan modal disetor - bersih ');
$objPHPExcel->getActiveSheet()->setCellValue('B76', 'Modal Lainnya');
$objPHPExcel->getActiveSheet()->setCellValue('B77', 'Laba (Rugi) yang belum direalidasi atas perubahan');

$objPHPExcel->getActiveSheet()->setCellValue('B78', 'Nilai wajar efek tersedia untuk dijual - netto');
$objPHPExcel->getActiveSheet()->setCellValue('B79', 'Telah ditentukan penggunaannya');
$objPHPExcel->getActiveSheet()->setCellValue('B80', 'Belum ditentukan penggunaannya');
$objPHPExcel->getActiveSheet()->setCellValue('B81', 'JUMLAH EKUITAS ');
$objPHPExcel->getActiveSheet()->setCellValue('B82', 'JUMLAH LIABILTAS DAN EKUITAS ');
$objPHPExcel->getActiveSheet()->setCellValue('A86', 'Lihat catatan atas laporan keuangan yang merupakan bagian tak terpisahkan ');
$objPHPExcel->getActiveSheet()->setCellValue('A87', 'dari laporan keuangan secara keseluruhan ');

$objPHPExcel->getActiveSheet()->setCellValue('A100', 'PT BANK MNC INTERNASIONAL TBK.');
$objPHPExcel->getActiveSheet()->setCellValue('A101', "LAPORAN LABA RUGI KOMPREHENSIF");
$objPHPExcel->getActiveSheet()->setCellValue('A102', "Untuk periode yang berakhir pada tanggal $tanggal_awal");
$objPHPExcel->getActiveSheet()->setCellValue('A103', 'Disajikan Dalam Jutaan Rupiah Kecuali Dinyatakan yang lain');

$objPHPExcel->getActiveSheet()->setCellValue('B106', 'PENDAPATAN DAN BEBAN OPERASIONAL');
$objPHPExcel->getActiveSheet()->setCellValue('C106', $tanggal_awal);
$objPHPExcel->getActiveSheet()->setCellValue('B107', "Pendapatan Bunga");
$objPHPExcel->getActiveSheet()->setCellValue('B108', "Beban Bunga");
$objPHPExcel->getActiveSheet()->setCellValue('B109', 'Pendapatan Bunga Bersih ');
$objPHPExcel->getActiveSheet()->setCellValue('B110', 'Pendapatan (Beban) Operasional Lainnya ');
$objPHPExcel->getActiveSheet()->setCellValue('B111', "Pendapatan Operasional Lainnya :");
$objPHPExcel->getActiveSheet()->setCellValue('B112', "Keuntungan penjualan efek efek yang");
$objPHPExcel->getActiveSheet()->setCellValue('B113', 'diperdagangkan dan investasi keuangan bersih');
$objPHPExcel->getActiveSheet()->setCellValue('B114', "Pendapatan Provisi dan Komisi ");
$objPHPExcel->getActiveSheet()->setCellValue('B115', "Keuntungan dari transaksi mata uang");
$objPHPExcel->getActiveSheet()->setCellValue('B116', '     asing - bersih');
$objPHPExcel->getActiveSheet()->setCellValue('B117', "Peningkatan Nilai wajar (MTM) Aset Keuangan Lainnya");
$objPHPExcel->getActiveSheet()->setCellValue('B118', "Pendapatan Lain Lain");
$objPHPExcel->getActiveSheet()->setCellValue('B120', 'Jumlah Pendapatan Operasi Lainnya');
$objPHPExcel->getActiveSheet()->setCellValue('B121', 'Jumlah Pendapatan Operasional Bersih');
$objPHPExcel->getActiveSheet()->setCellValue('B123', 'Beban Operasional lainnya :');
$objPHPExcel->getActiveSheet()->setCellValue('B124', 'Beban penyisihan kerugian penurunan nilai');
$objPHPExcel->getActiveSheet()->setCellValue('B125', '     atas aset keuangan dan aset non keuangan');
$objPHPExcel->getActiveSheet()->setCellValue('B126', 'Penurunan nilai wajar (MTM) Surat Berharga');
$objPHPExcel->getActiveSheet()->setCellValue('B128', 'Umum dan administrasi');
$objPHPExcel->getActiveSheet()->setCellValue('B129', 'Tenaga Kerja');
$objPHPExcel->getActiveSheet()->setCellValue('B130', 'Jumlah Beban Operasional lainnya');
$objPHPExcel->getActiveSheet()->setCellValue('B131', 'Pendapatan (Rugi) Operasional Bersih
');  
$objPHPExcel->getActiveSheet()->setCellValue('B133', 'PENDAPATAN (BEBAN) NON OPERASIONAL');
$objPHPExcel->getActiveSheet()->setCellValue('B134', 'Keuntungan Penjualan Aset Tetap Bersih
');
$objPHPExcel->getActiveSheet()->setCellValue('B135', 'Keuntungan / (Kerugian) Penjualan AYDA');
$objPHPExcel->getActiveSheet()->setCellValue('B136', 'Lainnya Bersih');
$objPHPExcel->getActiveSheet()->setCellValue('B137', 'Pendapatan Non Operasional');  

$objPHPExcel->getActiveSheet()->setCellValue('B139', 'LABA (RUGI) SEBELUM TAKSIRAN PAJAK PENGHASILAN');
$objPHPExcel->getActiveSheet()->setCellValue('B140', 'TAKSIRAN BEBAN PAJAK PENGHASILAN'); 
$objPHPExcel->getActiveSheet()->setCellValue('B141', 'LABA (RUGI) BERSIH');



$baris=10;
for ($i=1; $i <=5 ; $i++) { 
$objPHPExcel->getActiveSheet()->setCellValue("C$baris", floatval($nominal_neraca["$i"]));
$baris++;
}

$objPHPExcel->getActiveSheet()->setCellValue("C15", floatval($nominal_neraca["8"]));
$objPHPExcel->getActiveSheet()->setCellValue("C16", "=SUM(C14:C15)");
$objPHPExcel->getActiveSheet()->setCellValue("C17", floatval($nominal_neraca["6"]));
$objPHPExcel->getActiveSheet()->setCellValue("C18", floatval($nominal_neraca["7"]));
$objPHPExcel->getActiveSheet()->setCellValue("C19", 0);
$objPHPExcel->getActiveSheet()->setCellValue("C20", floatval($nominal_neraca["7"]));
$objPHPExcel->getActiveSheet()->setCellValue("C21", "=SUM(C19:C20)");
$objPHPExcel->getActiveSheet()->setCellValue("C22", floatval($nominal_neraca["8"]));
$objPHPExcel->getActiveSheet()->setCellValue("C23", floatval($nominal_neraca["9"]));
$objPHPExcel->getActiveSheet()->setCellValue("C24", 0);
$objPHPExcel->getActiveSheet()->setCellValue("C25", floatval($nominal_neraca["11"]));
$objPHPExcel->getActiveSheet()->setCellValue("C26", floatval($nominal_neraca["14"]));
$objPHPExcel->getActiveSheet()->setCellValue("C27", 0);
$objPHPExcel->getActiveSheet()->setCellValue("C28", 0);
$objPHPExcel->getActiveSheet()->setCellValue("C29", floatval($nominal_neraca["13"]));
$objPHPExcel->getActiveSheet()->setCellValue("C30", 0);


$objPHPExcel->getActiveSheet()->setCellValue("C56", floatval($nominal_neraca["15"]));
$objPHPExcel->getActiveSheet()->setCellValue("C57", floatval($nominal_neraca["16"]));
$objPHPExcel->getActiveSheet()->setCellValue("C58", 0);
$objPHPExcel->getActiveSheet()->setCellValue("C59", floatval($nominal_neraca["16"]));
$objPHPExcel->getActiveSheet()->setCellValue("C60", "=SUM(C58:C59)");
$objPHPExcel->getActiveSheet()->setCellValue("C61", floatval($nominal_neraca["17"]));
$objPHPExcel->getActiveSheet()->setCellValue("C62", floatval($nominal_neraca["21"]));
$objPHPExcel->getActiveSheet()->setCellValue("C63", floatval($nominal_neraca["18"]));
$objPHPExcel->getActiveSheet()->setCellValue("C64", floatval($nominal_neraca["20"]));
$objPHPExcel->getActiveSheet()->setCellValue("C65", floatval($nominal_neraca["23"]));
$objPHPExcel->getActiveSheet()->setCellValue("C66", 0);
$objPHPExcel->getActiveSheet()->setCellValue("C67", floatval($nominal_neraca["24"]));
$objPHPExcel->getActiveSheet()->setCellValue("C68", 0);
$objPHPExcel->getActiveSheet()->setCellValue("C69", 0);
$objPHPExcel->getActiveSheet()->setCellValue("C70", floatval($nominal_neraca["26"]));
$objPHPExcel->getActiveSheet()->setCellValue("C71", floatval($nominal_neraca["28"]));
$objPHPExcel->getActiveSheet()->setCellValue("C72", 0);
$objPHPExcel->getActiveSheet()->setCellValue("C73", 0);
$objPHPExcel->getActiveSheet()->setCellValue("C74", floatval($nominal_neraca["30"]));
$objPHPExcel->getActiveSheet()->setCellValue("C75", floatval($nominal_neraca["31"]));
$objPHPExcel->getActiveSheet()->setCellValue("C76", 0);
$objPHPExcel->getActiveSheet()->setCellValue("C77", 0);
$objPHPExcel->getActiveSheet()->setCellValue("C78", 0);
$objPHPExcel->getActiveSheet()->setCellValue("C79", 0);
$objPHPExcel->getActiveSheet()->setCellValue("C80", 0);
$objPHPExcel->getActiveSheet()->setCellValue("C81", 0);
$objPHPExcel->getActiveSheet()->setCellValue("C82", 0);

$objPHPExcel->getActiveSheet()->setCellValue("C107", 0);
$objPHPExcel->getActiveSheet()->setCellValue("C108", 0);
$objPHPExcel->getActiveSheet()->setCellValue("C109", 0);
$objPHPExcel->getActiveSheet()->setCellValue("C110", 0);
$objPHPExcel->getActiveSheet()->setCellValue("C111", 0);
$objPHPExcel->getActiveSheet()->setCellValue("C112", 0);
$objPHPExcel->getActiveSheet()->setCellValue("C113", 0);
$objPHPExcel->getActiveSheet()->setCellValue("C114", 0);
$objPHPExcel->getActiveSheet()->setCellValue("C115", 0);
$objPHPExcel->getActiveSheet()->setCellValue("C116", 0);
$objPHPExcel->getActiveSheet()->setCellValue("C117", 0);
$objPHPExcel->getActiveSheet()->setCellValue("C118", 0);
$objPHPExcel->getActiveSheet()->setCellValue("C119", 0);
$objPHPExcel->getActiveSheet()->setCellValue("C120", 0);
$objPHPExcel->getActiveSheet()->setCellValue("C121", 0);
$objPHPExcel->getActiveSheet()->setCellValue("C122", 0);
$objPHPExcel->getActiveSheet()->setCellValue("C123", 0);
$objPHPExcel->getActiveSheet()->setCellValue("C124", 0);
$objPHPExcel->getActiveSheet()->setCellValue("C125", 0);
$objPHPExcel->getActiveSheet()->setCellValue("C126", 0);
$objPHPExcel->getActiveSheet()->setCellValue("C127", 0);
$objPHPExcel->getActiveSheet()->setCellValue("C128", floatval($nominal_neraca["46"]));
$objPHPExcel->getActiveSheet()->setCellValue("C129", floatval($nominal_neraca["47"]));
$objPHPExcel->getActiveSheet()->setCellValue("C130", 0);
$objPHPExcel->getActiveSheet()->setCellValue("C131", 0);
$objPHPExcel->getActiveSheet()->setCellValue("C132", 0);
$objPHPExcel->getActiveSheet()->setCellValue("C133", 0);
$objPHPExcel->getActiveSheet()->setCellValue("C134", floatval($nominal_neraca["51"]));
$objPHPExcel->getActiveSheet()->setCellValue("C135", 0);
$objPHPExcel->getActiveSheet()->setCellValue("C136", floatval($nominal_neraca["52"]));
$objPHPExcel->getActiveSheet()->setCellValue("C137", 0);
$objPHPExcel->getActiveSheet()->setCellValue("C138", 0);
$objPHPExcel->getActiveSheet()->setCellValue("C139", 0);
$objPHPExcel->getActiveSheet()->setCellValue("C140", 0);
$objPHPExcel->getActiveSheet()->setCellValue("C141", 0);


$objPHPExcel->getActiveSheet()->getStyle("C10:C30")->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle("C56:C82")->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');
$objPHPExcel->getActiveSheet()->getStyle("C107:C141")->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');



$objPHPExcel->getActiveSheet()->setTitle('UNTUK_MNC');

$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
$objWriter->save("download/Report_LONGFORM_".$label_tgl."_".$file_eksport.".xls");


// LOAD FROM EXCEL FILE
$objPHPExcel = PHPExcel_IOFactory::load("download/Report_LONGFORM_".$label_tgl."_".$file_eksport.".xls");
$objWorksheet = $objPHPExcel->getActiveSheet();
$objPHPExcel->setActiveSheetIndex(0);

;
?>



<div class="portlet box grey-cascade" id="flash-report" >
                        <div class="portlet-title">
                            <div class="caption">
                                <i class="fa fa-book"></i> LAPORAN LONGFORM
                            </div>
                            <div class="tools">
                                <a href="javascript:;" class="collapse">
                                </a>

                                <a href="#portlet-config" data-toggle="modal" class="config">
                                </a>
                            </div>
                        </div>
                        <div class="portlet-body">
                            <h4 ><b>PT. Bank MNC Internasional .Tbk</b></h4>
                            <div class="tabbable-line">
                                <ul class="nav nav-tabs ">
                                    <li class="active">
                                        <a href="#tab_15_1" data-toggle="tab">
                                        Neraca Laba Rugi </a>
                                    </li>
                                    <li>
                                        <a href="#tab_15_2" data-toggle="tab">
                                        Untuk MNC </a>
                                    </li>
                                  
                                    
                                </ul>
                                <div class="tab-content">
                                    <div class="tab-pane active" id="tab_15_1">
                                     
                                            <h5>
                                            <b><div class="pull-right" style="font-size:12px">
<a href="<?php echo "http://".$_SERVER['HTTP_HOST']."/fincon_dev/module/report/"."download/Report_LONGFORM_".$label_tgl."_".$file_eksport.".xls";?>" class="btn btn-sm green"> Download Excel <i class="fa fa-arrow-circle-o-down"></i> </a> <br><br> </div> </b></h5>

</br>
</br>
    <div class="alert alert-info" >
                                        <button class="close" data-close="alert"></button>
                                       <b> Neraca Laba Rugi</b>
                                    </div>                                  
                                        
                                        <p>
                                        
                                        <div class="table-scrollable">
                                            <table class="table table-striped table-bordered table-advance table-hover"  width="100%">
                                                <thead>
                                               
                                                <tr class="active">
                                                <td width="50%" align="left"><b><?php echo $objPHPExcel->getActiveSheet()->getCell("B7")->getValue(); ?></b></td>
                                                <td width="25%" align="center" ><b><?php echo $objPHPExcel->getActiveSheet()->getCell("C7")->getValue(); ?></b></td>
                                                <td width="25%" align="center" ><b><?php echo $objPHPExcel->getActiveSheet()->getCell("D7")->getValue(); ?></b></td>
                                                </tr>
                                                </thead>
                                                <tbody>
                                                
                                              

                                                <?php
                                                



                                                for ($i=9; $i <=26 ; $i++) { 
                                                ?>

                                                <tr>
                                                <td  style="font-size:12px"><?php echo $objPHPExcel->getActiveSheet()->getCell("B$i")->getValue(); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <?php

}
?>
                                                <tr class="active">
                                                <td width="100%" colspan="3" align="center"><b>Neraca </b></td>
                                                </tr>
<?php
for ($i=45; $i <=75 ; $i++) { 

?>
                                               

                                                <tr>
                                                <td  style="font-size:12px"><?php echo $objPHPExcel->getActiveSheet()->getCell("B$i")->getValue(); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
<?php

}
?>

                                                <tr class="active">
                                                <td width="100%" colspan="3" align="center"><b>Laporan Rugi - Laba</b></td>
                                               
                                                </tr>

<?php
for ($i=91; $i <=134 ; $i++) { 
?>                                              
                                                <tr>
                                                <td  style="font-size:12px"><?php echo $objPHPExcel->getActiveSheet()->getCell("B$i")->getValue(); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>

<?php

}
 ?>                                              
                                                
     
                                                </tbody>
                                            </table>
                                        </div>
                                        </p>
                                    </div>
                                  
                                     <div class="tab-pane" id="tab_15_2">
                                     
                                        
<div class="alert alert-info" >
                                        <button class="close" data-close="alert"></button>
                                       <b> Laporan Longform untuk MNC </b>
                                    </div>  
                                      <?php
                                      $objPHPExcel->setActiveSheetIndex(1);
                                      ?>
                                        
                                        <p>
                                         <div class="table-scrollable">
                                            <table class="table table-striped table-bordered table-advance table-hover"  width="100%">
                                                <thead>
                                               
                                                <tr class="active">
                                                <td width="70%" align="left"><b><?php echo $objPHPExcel->getActiveSheet()->getCell("B7")->getValue(); ?></b></td>
                                                <td width="30%" align="center" ><b><?php echo $objPHPExcel->getActiveSheet()->getCell("C7")->getValue(); ?></b></td>
                                                
                                                </tr>
                                                </thead>
                                                <tbody>
                                                
                                              

                                                <?php
                                                



                                                for ($i=8; $i <=30 ; $i++) { 
                                                ?>

                                                <tr>
                                                <td  style="font-size:12px"><?php echo $objPHPExcel->getActiveSheet()->getCell("B$i")->getValue(); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                
                                                </tr>
                                                <?php

}
?>
                                                <tr class="active">
                                                <td width="100%" colspan="3" align="center"><b>Neraca </b></td>
                                                </tr>
<?php
for ($i=54; $i <=82 ; $i++) { 

?>
                                               

                                                <tr>
                                                <td  style="font-size:12px"><?php echo $objPHPExcel->getActiveSheet()->getCell("B$i")->getValue(); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                
                                                </tr>
<?php

}
?>

                                                <tr class="active">
                                                <td width="100%" colspan="3" align="center"><b>Laporan Rugi - Laba</b></td>
                                               
                                                </tr>

<?php
for ($i=106; $i <=141 ; $i++) { 
?>                                              
                                                <tr>
                                                <td  style="font-size:12px"><?php echo $objPHPExcel->getActiveSheet()->getCell("B$i")->getValue(); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                
                                                </tr>

<?php

}
 ?>                                              
                                                
     
                                                </tbody>
                                            </table>
                                        </div>
                                        
                                        </p>
                                    </div>

                                </div>
                            </div>
                            
                        </div>
                </div>

