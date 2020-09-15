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
logActivity("generate publikasi",date('Y_m_d_H_i_s'));
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
$kuartal=$_POST['kuartal'];
//echo $kuartal;
//die();
switch ($kuartal) {
        case '1':
        $tanggal=$tahun."-03-31";
        break;
        case '2':
        $tanggal=$tahun."-06-30";
        break;
        case '3':
        $tanggal=$tahun."-09-30";
        break;
        case '4':
        $tanggal=$tahun."-12-31";
        break;
     
}
$var_tgl=date('Y-m-d',strtotime($tanggal));




$query = " "; 


       $nama_debitur=array();
       $res=odbc_exec($connection2, $query);
while ($row=odbc_fetch_array($res)) {
      array_push($nama_debitur,$row['NamaLengkapDebitur']);
}

$query2 = "";

  $nama_debitur2=array();
       $res2=odbc_exec($connection2, $query2);
while ($row2=odbc_fetch_array($res2)) {
      array_push($nama_debitur2,$row2['NamaLengkapDebitur']);
}




// Create new PHPExcel object
$objPHPExcel = new PHPExcel();

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
$objPHPExcel->getActiveSheet()->getStyle('A4')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('A5')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('A7:F7')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('B9:F11')->applyFromArray($styleArrayFontBold);

$objPHPExcel->getActiveSheet()->getStyle('A4:E6')->applyFromArray($styleArrayAlignment1);

$objPHPExcel->getActiveSheet()->getStyle('B9')->applyFromArray($styleArrayAlignment1);
//$objPHPExcel->getActiveSheet()->getStyle('B9')->applyFromArray($styleArrayAlignment2);
$objPHPExcel->getActiveSheet()->getStyle('E9')->applyFromArray($styleArrayAlignment1);
$objPHPExcel->getActiveSheet()->getStyle('C9')->applyFromArray($styleArrayAlignment1);
$objPHPExcel->getActiveSheet()->getStyle('C11')->applyFromArray($styleArrayAlignment1);
$objPHPExcel->getActiveSheet()->getStyle('D11')->applyFromArray($styleArrayAlignment1);
$objPHPExcel->getActiveSheet()->getStyle('E11')->applyFromArray($styleArrayAlignment1);
$objPHPExcel->getActiveSheet()->getStyle('B21')->applyFromArray($styleArrayAlignment1);

//=======BORDER





//FILL COLOR
//$objPHPExcel->getActiveSheet()->getStyle('A10:E10')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('D3D3D3');



//DIMENSION D
$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(5);
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(50);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(30);
$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(30);
$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth(50);

// Create a first sheet, representing sales data
$objPHPExcel->setActiveSheetIndex(0);
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('A4:E4');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('A5:E5');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('A6:E6');

$objPHPExcel->setActiveSheetIndex(0)->mergeCells('B9:B11');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('C9:D10');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('E9:E10');

$objPHPExcel->getActiveSheet()->getStyle('C13:E21')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');


$objPHPExcel->getActiveSheet()->setCellValue('A4', 'LAPORAN REALISASI PEMBERIAN KREDIT ATAU PEMBIAYAAN UMKM MELALUI KERJASAMA POLA EXECUTING');
$objPHPExcel->getActiveSheet()->setCellValue('A5', "POSISI TRIWULAN  II  TAHUN $tahun ");
$objPHPExcel->getActiveSheet()->setCellValue('A7', 'A.');
$objPHPExcel->getActiveSheet()->setCellValue('B7', 'Laporan Baki Debet Kredit atau Pembiayaan UMKM Melalui Kerja Sama Pola Executing');
$objPHPExcel->getActiveSheet()->setCellValue('B9', 'PIHAK PENYALUR');

$objPHPExcel->getActiveSheet()->setCellValue('C9', 'Jenis Penggunaan ( Dalam Rupiah )');
$objPHPExcel->getActiveSheet()->setCellValue('C11', 'Modal Kerja');
$objPHPExcel->getActiveSheet()->setCellValue('D11', 'Investasi');

$objPHPExcel->getActiveSheet()->setCellValue('E9', 'TOTAL KREDIT / PEMBIAYAAN UMKM');
$objPHPExcel->getActiveSheet()->setCellValue('E11', '(Dalam Rupiah)');



$objPHPExcel->getActiveSheet()->setCellValue('B12', "BPR");





$number_row1=13;

$query_A=" ";


//echo $query_A;
//die();
        $result1=odbc_exec($connection2, $query_A);
        while ($row1=odbc_fetch_array($result1)) {
        $objPHPExcel->getActiveSheet()->setCellValue("B$number_row1", $row1['PihakPenyalur']);
        $objPHPExcel->getActiveSheet()->setCellValue("C$number_row1", $row1['ModalKerja']);
        
        $objPHPExcel->getActiveSheet()->setCellValue("E$number_row1", "=(C$number_row1+D$number_row1)");

    $number_row1++;
}



$no_total=$number_row1+1;


$objPHPExcel->getActiveSheet()->setCellValue("B$number_row1", 'TOTAL');

# TOTAL 
$objPHPExcel->getActiveSheet()->setCellValue("C$number_row1", "=SUM(C13:C".($number_row1-1).")");
$objPHPExcel->getActiveSheet()->setCellValue("D$number_row1", "=SUM(D13:D".($number_row1-1).")");
$objPHPExcel->getActiveSheet()->setCellValue("E$number_row1", "=SUM(E13:E".($number_row1-1).")");

// TITLE 
$objPHPExcel->getActiveSheet()->setTitle('LAPORAN A');

for ($i=13;$i<$number_row1;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('C'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        $objPHPExcel->getActiveSheet()->setCellValue('C'.$i, 0);
    }
}
for ($i=13;$i<$number_row1;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('D'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        $objPHPExcel->getActiveSheet()->setCellValue('D'.$i, 0);
    }
}

for ($i=13;$i<$number_row1;$i++) {
    //$kolom=array("C","D",);
    $colB = $objPHPExcel->getActiveSheet()->getCell('E'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        $objPHPExcel->getActiveSheet()->setCellValue('E'.$i, 0);
    }
}

$styleArrayBorder1 = array('borders' => array('allborders' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('argb' => ''),),),);
$styleArrayBorder2 = array('borders' => array('outline' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('argb' => ''),),),);

$objPHPExcel->getActiveSheet()->getStyle("B9:E$number_row1")->applyFromArray($styleArrayBorder1);


$objPHPExcel->getActiveSheet()->getStyle('A1:Z8')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle("A".($number_row1+1).":Z1000")->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('F1:Z1000')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('A1:A1000')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');



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

$objPHPExcel->getActiveSheet()->getStyle('C7:G15')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');

$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(5);
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(50);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(30);
$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(30);
$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth(30);
$objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth(30);
$objPHPExcel->getActiveSheet()->getColumnDimension('G')->setWidth(30);

$objPHPExcel->setActiveSheetIndex(1)->mergeCells('B4:B5');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('C4:F4');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('G4:G5');

$objPHPExcel->getActiveSheet()->getStyle('A2:F5')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('B15')->applyFromArray($styleArrayFontBold);




$objPHPExcel->getActiveSheet()->getStyle('A4:G5')->applyFromArray($styleArrayAlignment1);
$objPHPExcel->getActiveSheet()->getStyle('B15')->applyFromArray($styleArrayAlignment1);

//$objPHPExcel->getActiveSheet()->getStyle('B4:G15')->applyFromArray($styleArrayBorder1);


$objPHPExcel->getActiveSheet()->setCellValue('A2', 'B.');
$objPHPExcel->getActiveSheet()->setCellValue('B2', 'Laporan Kualitas Kredit atau Pembiayaan UMKM Melalui Kerja Sama Pola Executing.');


$objPHPExcel->getActiveSheet()->setCellValue('B4', 'PIHAK PENYALUR');
$objPHPExcel->getActiveSheet()->setCellValue('C4', 'KUALITAS ( Dalam Rupiah )');
$objPHPExcel->getActiveSheet()->setCellValue('G4', 'TOTAL');

$objPHPExcel->getActiveSheet()->setCellValue('C5', 'LANCAR ');
$objPHPExcel->getActiveSheet()->setCellValue('D5', 'KURANG LANCAR');
$objPHPExcel->getActiveSheet()->setCellValue('E5', 'DIRAGUKAN');
$objPHPExcel->getActiveSheet()->setCellValue('F5', 'MACET');





$objPHPExcel->getActiveSheet()->setCellValue('B6', "BPR");





// LANCAR
$number_row2=7;

        $result_lancar=odbc_exec($connection2, $query_lancar);
        while ($row_lancar=odbc_fetch_array($result_lancar)) {
        $objPHPExcel->getActiveSheet()->setCellValue("B$number_row2", $row_lancar['PihakPenyalur']);
        $objPHPExcel->getActiveSheet()->setCellValue("C$number_row2", $row_lancar['lancar']);     
        $objPHPExcel->getActiveSheet()->setCellValue("G$number_row2", "=C$number_row2+D$number_row2+E$number_row2+F$number_row2");



    $number_row2++;
}
// KURANG LANCAR

$number_row3=7;

        $result_krg_lancar=odbc_exec($connection2, $query_krg_lancar);
        while ($row_krg_lancar=odbc_fetch_array($result_krg_lancar)) {
        //$objPHPExcel->getActiveSheet()->setCellValue("B$number_row2", $row_lancar['PihakPenyalur']);
        $objPHPExcel->getActiveSheet()->setCellValue("D$number_row3", $row_krg_lancar['krg_lancar']);     
        //$objPHPExcel->getActiveSheet()->setCellValue("G$number_row2", "=C$number_row2+D$number_row2+E$number_row2+F$number_row2");
    $number_row3++;
}

// DIRAGUKAN

$number_row4=7;

        $result_diragukan=odbc_exec($connection2, $query_diragukan);
        while ($row_diragukan=odbc_fetch_array($result_diragukan)) {
        //$objPHPExcel->getActiveSheet()->setCellValue("B$number_row2", $row_lancar['PihakPenyalur']);
        $objPHPExcel->getActiveSheet()->setCellValue("E$number_row4", $row_diragukan['diragukan']);     
        //$objPHPExcel->getActiveSheet()->setCellValue("G$number_row2", "=C$number_row2+D$number_row2+E$number_row2+F$number_row2");
    $number_row4++;
}

// MACET

$number_row5=7;

        $result_macet=odbc_exec($connection2, $query_macet);
        while ($row_macet=odbc_fetch_array($result_macet)) {
        //$objPHPExcel->getActiveSheet()->setCellValue("B$number_row2", $row_lancar['PihakPenyalur']);
        $objPHPExcel->getActiveSheet()->setCellValue("F$number_row5", $row_macet['macet']);     
        //$objPHPExcel->getActiveSheet()->setCellValue("G$number_row2", "=C$number_row2+D$number_row2+E$number_row2+F$number_row2");
    $number_row5++;
}

$no_total2=$number_row2+1;

$objPHPExcel->getActiveSheet()->setCellValue("C$number_row2", "=SUM(C7:C".($number_row2-1).")");
$objPHPExcel->getActiveSheet()->setCellValue("D$number_row2", "=SUM(D7:D".($number_row2-1).")");
$objPHPExcel->getActiveSheet()->setCellValue("E$number_row2", "=SUM(E7:E".($number_row2-1).")");
$objPHPExcel->getActiveSheet()->setCellValue("F$number_row2", "=SUM(F7:F".($number_row2-1).")");
$objPHPExcel->getActiveSheet()->setCellValue("G$number_row2", "=SUM(G7:G".($number_row2-1).")");

$objPHPExcel->getActiveSheet()->setCellValue("B$number_row2", "TOTAL");




for ($i=7;$i<$number_row2;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('C'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        $objPHPExcel->getActiveSheet()->setCellValue('C'.$i, 0);
    }
}
for ($i=7;$i<$number_row2;$i++) {
    $colB = $objPHPExcel->getActiveSheet()->getCell('D'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        $objPHPExcel->getActiveSheet()->setCellValue('D'.$i, 0);
    }
}

for ($i=7;$i<$number_row2;$i++) {
    //$kolom=array("C","D",);
    $colB = $objPHPExcel->getActiveSheet()->getCell('E'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        $objPHPExcel->getActiveSheet()->setCellValue('E'.$i, 0);
    }
}
for ($i=7;$i<$number_row2;$i++) {
    //$kolom=array("C","D",);
    $colB = $objPHPExcel->getActiveSheet()->getCell('F'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        $objPHPExcel->getActiveSheet()->setCellValue('F'.$i, 0);
    }
}
for ($i=7;$i<$number_row2;$i++) {
    //$kolom=array("C","D",);
    $colB = $objPHPExcel->getActiveSheet()->getCell('G'.$i)->getValue();
    if ($colB == NULL || $colB == '') {
        $objPHPExcel->getActiveSheet()->setCellValue('G'.$i, 0);
    }
}

$objPHPExcel->getActiveSheet()->getStyle("B4:G$number_row2")->applyFromArray($styleArrayBorder1);

$objPHPExcel->getActiveSheet()->getStyle('A1:Z3')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle("A".($number_row2+1).":Z1000")->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('H1:Z1000')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');
$objPHPExcel->getActiveSheet()->getStyle('A1:A1000')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');


$objPHPExcel->getActiveSheet()->setTitle('LAPORAN B');

$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
$objWriter->save("download/Report_PUBLIKASI_".$label_tgl."_".$file_eksport.".xls");


// LOAD FROM EXCEL FILE
$objPHPExcel = PHPExcel_IOFactory::load("download/Report_PUBLIKASI_".$label_tgl."_".$file_eksport.".xls");
$objWorksheet = $objPHPExcel->getActiveSheet();
$objPHPExcel->setActiveSheetIndex(0);

;
?>



<div class="portlet box grey-cascade" id="flash-report" >
                        <div class="portlet-title">
                            <div class="caption">
                                <i class="fa fa-book"></i> LAPORAN UMKM 
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
                                        LAPORAN A </a>
                                    </li>
                                    <li>
                                        <a href="#tab_15_2" data-toggle="tab">
                                        LAPORAN B </a>
                                    </li>
                                  
                                    
                                </ul>
                                <div class="tab-content">
                                    <div class="tab-pane active" id="tab_15_1">
                                     
                                            <h5>
                                            <b><div class="pull-right" style="font-size:12px">
<a href="<?php echo "http://".$_SERVER['HTTP_HOST']."/fincon_dev/module/report/"."download/Report_UMKM_".$label_tgl."_".$file_eksport.".xls";?>" class="btn btn-sm green"> Download Excel <i class="fa fa-arrow-circle-o-down"></i> </a> <br><br> </div> </b></h5>

</br>
</br>
    <div class="alert alert-info" >
                                        <button class="close" data-close="alert"></button>
                                       <b> A. Laporan Baki Debet Kredit atau Pembiayaan UMKM Melalui Kerja Sama Pola Executing</b>
                                    </div>                                  
                                        
                                        <p>
                                        
                                        <div class="table-scrollable">
                                            <table class="table table-striped table-bordered table-advance table-hover"  width="100%">
                                                <thead>
                                               <tr class="active">
                                                <td width="30%" align="left"><b></b></td>
                                                <td width="50%" align="center" colspan="2"><b>Jenis Penggunaan (Dalam Rupiah) </b></td>
                                                <td width="20%" align="center"><b>Total Kredit / Pembiayaan UMKM </b></td>
                                               
                                                </tr>
                                                <tr class="active">
                                                <td width="30%" align="center"><b>PIHAK PENYALUR </b></td>
                                                <td width="25%" align="center" ><b>Modal Kerja </b></td>
                                                <td width="25%" align="center" ><b>Investasi </b></td>
                                                <td width="20%" align="center"><b> (Dalam Rupiah) </b></td>
                                                </tr>

                                                </thead>
                                                <tbody>
                                                
                                                <tr>
                                                <td  style="font-size:12px" align="left" >BPR </td>
                                                <td  style="font-size:12px" ></td>
                                                <td  style="font-size:12px"></td>
                                                <td  style="font-size:12px"></td>
                                                </tr>

                                                <?php
                                                $number_dash1=13;
                                                 foreach ($nama_debitur as $val) {
                                                    ?>
                                                <tr>
                                                <td  style="font-size:12px"><?php echo $objPHPExcel->getActiveSheet()->getCell("B$number_dash1")->getValue(); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C$number_dash1")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D$number_dash1")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E$number_dash1")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>

<?php       
$number_dash1++;
}


                                                ?>
                                                
                                                <tr>
                                                <td  style="font-size:12px"><?php echo $objPHPExcel->getActiveSheet()->getCell("B$number_dash1")->getValue(); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C$number_dash1")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D$number_dash1")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  style="font-size:12px" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E$number_dash1")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>

                                                
                                                </tbody>
                                            </table>
                                        </div>
                                        </p>
                                    </div>
                                  
                                     <div class="tab-pane" id="tab_15_2">
                                     
                                        
<div class="alert alert-info" >
                                        <button class="close" data-close="alert"></button>
                                       <b> B. Laporan Kualitas Kredit atau Pembiayaan UMKM Melalui Kerja Sama Pola Executing.</b>
                                    </div>  

                                      <?php
                                      $objPHPExcel->setActiveSheetIndex(1);
                                      ?>
                                        
                                        <p>
                                        <div class="table-scrollable">
                                            <table class="table table-striped table-bordered table-advance table-hover"  width="100%">
                                                <thead>
                                                <tr class="active">
                                                <td width="25%" align="left"><b></b></td>
                                                <td width="60%" align="center" colspan="4"><b>KUALITAS (Dalam Rupiah) </b></td>
                                                <td width="15%" align="center"><b>Total  </b></td>
                                               
                                                </tr>
                                                <tr class="active">
                                                <td width="25%" align="center"><b>PIHAK PENYALUR </b></td>
                                                <td width="15%" align="center" ><b>Lancar </b></td>
                                                <td width="15%" align="center" ><b>Kurang Lancar </b></td>
                                                <td width="15%" align="center" ><b>Diragukan </b></td>
                                                <td width="15%" align="center" ><b>Macet </b></td>
                                                <td width="15%" align="center"><b></b></td>
                                                </tr>

                                                </thead>
                                                <tbody>
                                                
                                                <tr>
                                                <td  align="left">BPR </td>
                                                <td  align="center" ></td>
                                                <td  align="center" ></td>
                                                <td  align="center" ></td>
                                                <td  align="center" ></td>
                                                <td  align="center"></td>
                                                </tr>
<?php
$number_dash2=7;
foreach ($nama_debitur2 as $val2) {
?>
                                                <tr>
                                                <td  align="left"><?php echo $objPHPExcel->getActiveSheet()->getCell("B$number_dash2")->getValue(); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C$number_dash2")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D$number_dash2")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E$number_dash2")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F$number_dash2")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G$number_dash2")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
<?php
$number_dash2++;
}

?>
                                                <tr>
                                                <td  align="left"><?php echo $objPHPExcel->getActiveSheet()->getCell("B$number_dash2")->getValue(); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("C$number_dash2")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("D$number_dash2")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E$number_dash2")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("F$number_dash2")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td  align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("G$number_dash2")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                
                                                </tbody>
                                            </table>
                                        </div>
                                        
                                        </p>
                                    </div>

                                </div>
                            </div>
                            
                        </div>
                </div>

