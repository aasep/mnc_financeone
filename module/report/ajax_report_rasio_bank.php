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
logActivity("generate KPMM",date('Y_m_d_H_i_s'));

######## POST DATE ##############
$tanggal=$_POST['tanggal']; 

$curr_tgl=date('Y-m-d',strtotime($tanggal));
$end_curr_tgl=date('Y-m-t',strtotime($tanggal));
$label_tgl=date('d-M-y',strtotime($tanggal));


$label_txtfile=date('Ymd',strtotime($tanggal));
$tanggal_header=date('dmY',strtotime($tanggal));


$day=date('d',strtotime($tanggal));
$mon=date('M',strtotime($tanggal));
$year=date('y',strtotime($tanggal));

$mon_modal=date('n',strtotime($tanggal));
$year_modal=date('Y',strtotime($tanggal));

$label_tgl=$day."-".$mon."-".$year; // tanggal terpilih
$label_bln=$mon."-".$year; // Bulan terpilih


$var_tabel=date('Ymd',strtotime($tanggal));

// echo "TESST..........!";
// die();

/*
##### QUERY KPMM ############## 
$query=" SELECT SUM (Nilai)/1000000 AS Jumlah_Nominal FROM (
SELECT a.kodegl,SUM(a.Nominal) AS Nilai FROM DM_Journal a WITH (NOLOCK)
JOIN Referensi_GL_02 b ON b.GLNO = a.KodeGL AND b.PRODNO = a.KodeProduct
JOIN Referensi_KPMM c ON c.KPMM_Level_3 = b.KPMM_Level_3
WHERE a.DataDate='$curr_tgl'  ";



$q_add2="  GROUP BY a.kodegl ,b.KPMM_Level_3 )AS tabel1 ";

 $curr_mon=date('n',strtotime($tanggal));
 $curr_year=date('Y',strtotime($tanggal));
# +++++++++++ QUERY ATMR ++++++++++++++
$query_atmr= " select * from Master_ATMR WHERE  Month(DataDate)='$curr_mon' and Year(DataDate)='$curr_year'  ";
//echo $query_atmr;
//die();

$result_atmr=odbc_exec($connection2, $query_atmr);
$rowAtmr=odbc_fetch_array($result_atmr);
$atmr_kredit=$rowAtmr['ATMR_Kredit'];
$atmr_pasar=$rowAtmr['ATMR_Pasar'];
$atmr_operasional=$rowAtmr['ATMR_Operasional'];

$q_add=" and b.KPMM_Level_3='KPMM205000001' ";
$result=odbc_exec($connection2, $query.$q_add.$q_add2);
$row=odbc_fetch_array($result);
$m11=abs($row['Jumlah_Nominal']);


*/



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



$objPHPExcel->getActiveSheet()->getStyle('A1:H30')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');



$objPHPExcel->getActiveSheet()->getStyle('B2:F2')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('F2:F22')->applyFromArray($styleArrayFontBold);



$objPHPExcel->getActiveSheet()->getStyle('B2:F2')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

$objPHPExcel->getActiveSheet()->getStyle('B2:D23')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('B2:D23')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
// $objPHPExcel->getActiveSheet()->getStyle('M4')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);

//=======BORDER

$styleArrayBorder1 = array('borders' => array('allborders' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('argb' => ''),),),);
$styleArrayBorder2 = array('borders' => array('outline' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('argb' => ''),),),);



$objPHPExcel->setActiveSheetIndex(0);
$objPHPExcel->setActiveSheetIndex(0)->mergeCells("B4:B5");
$objPHPExcel->setActiveSheetIndex(0)->mergeCells("C4:C5");
$objPHPExcel->setActiveSheetIndex(0)->mergeCells("F4:F5");

$objPHPExcel->setActiveSheetIndex(0)->mergeCells("B6:B7");
$objPHPExcel->setActiveSheetIndex(0)->mergeCells("C6:C7");
$objPHPExcel->setActiveSheetIndex(0)->mergeCells("F6:F7");

$objPHPExcel->setActiveSheetIndex(0)->mergeCells("B8:B9");
$objPHPExcel->setActiveSheetIndex(0)->mergeCells("C8:C9");
$objPHPExcel->setActiveSheetIndex(0)->mergeCells("F8:F9");

$objPHPExcel->setActiveSheetIndex(0)->mergeCells("B10:B11");
$objPHPExcel->setActiveSheetIndex(0)->mergeCells("C10:C11");
$objPHPExcel->setActiveSheetIndex(0)->mergeCells("F10:F11");

$objPHPExcel->setActiveSheetIndex(0)->mergeCells("B12:B13");
$objPHPExcel->setActiveSheetIndex(0)->mergeCells("C12:C13");
$objPHPExcel->setActiveSheetIndex(0)->mergeCells("F12:F13");

$objPHPExcel->setActiveSheetIndex(0)->mergeCells("B14:B15");
$objPHPExcel->setActiveSheetIndex(0)->mergeCells("C14:C15");
$objPHPExcel->setActiveSheetIndex(0)->mergeCells("F14:F15");

$objPHPExcel->setActiveSheetIndex(0)->mergeCells("B16:B17");
$objPHPExcel->setActiveSheetIndex(0)->mergeCells("C16:C17");
$objPHPExcel->setActiveSheetIndex(0)->mergeCells("F16:F17");

$objPHPExcel->setActiveSheetIndex(0)->mergeCells("B18:B19");
$objPHPExcel->setActiveSheetIndex(0)->mergeCells("C18:C19");
$objPHPExcel->setActiveSheetIndex(0)->mergeCells("F18:F19");

$objPHPExcel->setActiveSheetIndex(0)->mergeCells("B20:B21");
$objPHPExcel->setActiveSheetIndex(0)->mergeCells("C20:C21");
$objPHPExcel->setActiveSheetIndex(0)->mergeCells("F20:F21");

$objPHPExcel->setActiveSheetIndex(0)->mergeCells("B22:B23");
$objPHPExcel->setActiveSheetIndex(0)->mergeCells("C22:C23");
$objPHPExcel->setActiveSheetIndex(0)->mergeCells("F22:F23");

// for ($i=4; $i<=23 ; $i+2) { 
//   $x=$i+1;
// $objPHPExcel->setActiveSheetIndex(0)->mergeCells("B$i:B$x");
// $objPHPExcel->setActiveSheetIndex(0)->mergeCells("C$i:C$x");
// $objPHPExcel->setActiveSheetIndex(0)->mergeCells("F$i:F$x");
// }
$objPHPExcel->getActiveSheet()->getStyle('B2:F3')->applyFromArray($styleArrayBorder1);
$objPHPExcel->getActiveSheet()->getStyle('B2:C23')->applyFromArray($styleArrayBorder1);
$objPHPExcel->getActiveSheet()->getStyle('E2:F23')->applyFromArray($styleArrayBorder1);
$objPHPExcel->getActiveSheet()->getStyle("D4:D5")->applyFromArray($styleArrayBorder2);
$objPHPExcel->getActiveSheet()->getStyle("D6:D7")->applyFromArray($styleArrayBorder2);
$objPHPExcel->getActiveSheet()->getStyle("D8:D9")->applyFromArray($styleArrayBorder2);
$objPHPExcel->getActiveSheet()->getStyle("D10:D11")->applyFromArray($styleArrayBorder2);
$objPHPExcel->getActiveSheet()->getStyle("D12:D13")->applyFromArray($styleArrayBorder2);
$objPHPExcel->getActiveSheet()->getStyle("D14:D15")->applyFromArray($styleArrayBorder2);
$objPHPExcel->getActiveSheet()->getStyle("D16:D17")->applyFromArray($styleArrayBorder2);
$objPHPExcel->getActiveSheet()->getStyle("D18:D19")->applyFromArray($styleArrayBorder2);
$objPHPExcel->getActiveSheet()->getStyle("D20:D21")->applyFromArray($styleArrayBorder2);
$objPHPExcel->getActiveSheet()->getStyle("D22:D23")->applyFromArray($styleArrayBorder2);
 // for ($i=4; $i<=23 ; $i+2) { 
 //   $x=$i+1;
 // $objPHPExcel->getActiveSheet()->getStyle("D$i:D$x")->applyFromArray($styleArrayBorder2);
 // }

//$objPHPExcel->getActiveSheet()->getStyle("D$i:D$x")->applyFromArray($styleArrayBorder2);
// for ($i=9; $i<=75 ; $i++) { 
// $objPHPExcel->getActiveSheet()->getStyle("A$i:L$i")->applyFromArray($styleArrayBorder2);
// }


//DIMENSION D
$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(7);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(75);
$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(60);
$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth(15);

//$objPHPExcel->getActiveSheet()->getRowDimension(33)->setRowHeight(30);


// Create a first sheet, representing sales data







$objPHPExcel->getActiveSheet()->setCellValue('B2', 'No');
$objPHPExcel->getActiveSheet()->setCellValue('C2', 'Rasio (%)');
$objPHPExcel->getActiveSheet()->setCellValue('D2', "Formula");
$objPHPExcel->getActiveSheet()->setCellValue('E2', "NOMINAL");
$objPHPExcel->getActiveSheet()->setCellValue('F2', "%");

$objPHPExcel->getActiveSheet()->setCellValue('B3', '1');
$objPHPExcel->getActiveSheet()->setCellValue('C3', 'Kewajiban penyedian Modal Minimum (KPMM) *)');
$objPHPExcel->getActiveSheet()->setCellValue('D3', "CAR");
$objPHPExcel->getActiveSheet()->setCellValue('E3', "");
$objPHPExcel->getActiveSheet()->setCellValue('F3', "");

$objPHPExcel->getActiveSheet()->setCellValue('B4', '2');
$objPHPExcel->getActiveSheet()->setCellValue('C4', 'Aset produktif bermasalah dan aset non produktif bermasalah terhadap total aset produktif');
$objPHPExcel->getActiveSheet()->setCellValue('D4', "Aset produktif bermasalah + Aset non produksi bermasalah");
$objPHPExcel->getActiveSheet()->setCellValue('D5', "Total aset produktif + total aset non produktif");
$objPHPExcel->getActiveSheet()->setCellValue('E4', "");
$objPHPExcel->getActiveSheet()->setCellValue('E5', "");
$objPHPExcel->getActiveSheet()->setCellValue('F4', "");

$objPHPExcel->getActiveSheet()->setCellValue('B6', '3');
$objPHPExcel->getActiveSheet()->setCellValue('C6', 'Aset produktif bermasalah terhadap total aset produktif ');
$objPHPExcel->getActiveSheet()->setCellValue('D6', "Aset Produktif bermasalah (diluar transaksi rekening administratif)");
$objPHPExcel->getActiveSheet()->setCellValue('D7', "Total aset produktif (diluar transaksi rekening administratif)");
$objPHPExcel->getActiveSheet()->setCellValue('E6', "");
$objPHPExcel->getActiveSheet()->setCellValue('E7', "");
$objPHPExcel->getActiveSheet()->setCellValue('F6', "");

$objPHPExcel->getActiveSheet()->setCellValue('B8', '4');
$objPHPExcel->getActiveSheet()->setCellValue('C8', 'Cadangan kerugian penurunan nilai (CKPN) aset keuangan terhadap aset produktif');
$objPHPExcel->getActiveSheet()->setCellValue('D8', "CKPN aset Keuangan");
$objPHPExcel->getActiveSheet()->setCellValue('D9', "Total aset produktif ( diluar transaksi rekening administratif )");
$objPHPExcel->getActiveSheet()->setCellValue('E8', "");
$objPHPExcel->getActiveSheet()->setCellValue('E9', "");
$objPHPExcel->getActiveSheet()->setCellValue('F8', "");

$objPHPExcel->getActiveSheet()->setCellValue('B10', '5');
$objPHPExcel->getActiveSheet()->setCellValue('C10', 'NPL gross');
$objPHPExcel->getActiveSheet()->setCellValue('D10', "Kredit bermasalah");
$objPHPExcel->getActiveSheet()->setCellValue('D11', "Total kredit");
$objPHPExcel->getActiveSheet()->setCellValue('E10', "");
$objPHPExcel->getActiveSheet()->setCellValue('E11', "");
$objPHPExcel->getActiveSheet()->setCellValue('F10', "");

$objPHPExcel->getActiveSheet()->setCellValue('B12', '6');
$objPHPExcel->getActiveSheet()->setCellValue('C12', 'NPL net');
$objPHPExcel->getActiveSheet()->setCellValue('D12', "Kredit bermasalah - CKPN kredit");
$objPHPExcel->getActiveSheet()->setCellValue('D13', "Total kredit");
$objPHPExcel->getActiveSheet()->setCellValue('E12', "");
$objPHPExcel->getActiveSheet()->setCellValue('E13', "");
$objPHPExcel->getActiveSheet()->setCellValue('F12', "");

$objPHPExcel->getActiveSheet()->setCellValue('B14', '7');
$objPHPExcel->getActiveSheet()->setCellValue('C14', 'Return on Asset (ROA)');
$objPHPExcel->getActiveSheet()->setCellValue('D14', "Laba (Rugi) Sebelum Pajak (disetahunkan)");
$objPHPExcel->getActiveSheet()->setCellValue('D15', "Rata-rata Total Assets");
$objPHPExcel->getActiveSheet()->setCellValue('E14', "");
$objPHPExcel->getActiveSheet()->setCellValue('E15', "");
$objPHPExcel->getActiveSheet()->setCellValue('F14', "");

$objPHPExcel->getActiveSheet()->setCellValue('B16', '8');
$objPHPExcel->getActiveSheet()->setCellValue('C16', 'Return on Equity (ROE)');
$objPHPExcel->getActiveSheet()->setCellValue('D16', "Laba (Rugi) Setelah Pajak (disetahunkan)");
$objPHPExcel->getActiveSheet()->setCellValue('D17', "Rata-rata Equity (Tier 1)");
$objPHPExcel->getActiveSheet()->setCellValue('E16', "");
$objPHPExcel->getActiveSheet()->setCellValue('E17', "");
$objPHPExcel->getActiveSheet()->setCellValue('F16', "");

$objPHPExcel->getActiveSheet()->setCellValue('B18', '9');
$objPHPExcel->getActiveSheet()->setCellValue('C18', 'Net Interest Margin (NIM)');
$objPHPExcel->getActiveSheet()->setCellValue('D18', "Pend. Bunga Bersih (disetahunkan)");
$objPHPExcel->getActiveSheet()->setCellValue('D19', "Rata-rata Aktiva Prod.");
$objPHPExcel->getActiveSheet()->setCellValue('E18', "");
$objPHPExcel->getActiveSheet()->setCellValue('E19', "");
$objPHPExcel->getActiveSheet()->setCellValue('F18', "");

$objPHPExcel->getActiveSheet()->setCellValue('B20', '10');
$objPHPExcel->getActiveSheet()->setCellValue('C20', 'Biaya Operasional terhadap Pendapatan Operasional (BOPO)');
$objPHPExcel->getActiveSheet()->setCellValue('D20', "Total Beban Operasional");
$objPHPExcel->getActiveSheet()->setCellValue('D21', "Total Pend. Operasional");
$objPHPExcel->getActiveSheet()->setCellValue('E20', "");
$objPHPExcel->getActiveSheet()->setCellValue('E21', "");
$objPHPExcel->getActiveSheet()->setCellValue('F20', "");

$objPHPExcel->getActiveSheet()->setCellValue('B22', '11');
$objPHPExcel->getActiveSheet()->setCellValue('C22', 'Loan to Funding Ratio (LFR)');
$objPHPExcel->getActiveSheet()->setCellValue('D22', "Kredit");
$objPHPExcel->getActiveSheet()->setCellValue('D23', "Dana Pihak Ketiga ");
$objPHPExcel->getActiveSheet()->setCellValue('E22', "");
$objPHPExcel->getActiveSheet()->setCellValue('E23', "");
$objPHPExcel->getActiveSheet()->setCellValue('F22', "");

 
// $objPHPExcel->getActiveSheet()->getStyle('H77')->getNumberFormat()->applyFromArray( array( 'code' => PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00 ));     
//$objPHPExcel->getActiveSheet()->getStyle('A9:L9')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('483D8B');
//$objPHPExcel->getActiveSheet()->getStyle('A5:L7')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('DCDCDC');


//$objPHPExcel->getActiveSheet()->getStyle('C17:H22')->getNumberFormat()->setFormatCode('0.00');

 
 $objPHPExcel->getActiveSheet()->getStyle('E3:E23')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');


$objPHPExcel->getActiveSheet()->setTitle('RASIO');

$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
$objWriter->save("download/RASIO_".$label_tgl."_".$file_eksport.".xls");


// LOAD FROM EXCEL FILE
$objPHPExcel = PHPExcel_IOFactory::load("download/RASIO_".$label_tgl."_".$file_eksport.".xls");
$objWorksheet = $objPHPExcel->getActiveSheet();



?>

<div class="portlet box blue" id="flash-report" >
                        <div class="portlet-title">
                            <div class="caption">
                                <i class="fa fa-book"></i> List Laporan Rasio Bank  
                            </div>
                          
                        </div>
                        <div class="portlet-body">
                            <h4 ><b>PT Bank MNC Internasional, Tbk</b></h4>
                            
                            
                            <div class="tabbable-line">
                            
                                <div class="tab-content">
                                    <div class="tab-pane active" id="tab_15_1">
                                     
                                            <h5>
                                            <b>
												<div class="pull-right" style="font-size:12px">
													<a href="<?php echo "http://".$_SERVER['HTTP_HOST']."/fincon_dev/module/report/"."download/RASIO_".$label_tgl."_".$file_eksport.".xls";?>" class="btn btn-sm green"> Download Excel <i class="fa fa-arrow-circle-o-down"></i> 
													</a> 
												</div>
												<br>
												<br>
												<br>
												<br>
												<div class="pull-right" >(dalam jutaan rupiah)</div>
											</b> 

                                            

                                       
                                         
                                        <div class="table-scrollable">
                                            <table class="table table-striped table-bordered table-advance table-hover"  width="100%">
                                                <thead>
                                                <tr class="active">
                                                <td width="5%" align="center" ><b>No</b></td>
                                                <td width="35%" align="center" ><b> Rasio (%) </b></td>
                                                <td width="30%" align="center" ><b> Formula</b></td>
                                                <td width="15%" align="center" ><b> Nominal</b></td>
                                                <td width="15%" align="right" ><b> %</b></td>
                                                </tr>
                                             
                                                </thead>
                                                <tbody>

                                               <?php 

                                               for ($i=3; $i <= 23 ; $i++) { 
                                               	
                                                if ($i=='3'){   ?>
                                                <tr>
                                                <td align="center" width="5%"  > <b><?php echo $objPHPExcel->getActiveSheet()->getCell("B$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"');?></b></td>
                                                <td align="center" width="35%"  > <?php echo $objPHPExcel->getActiveSheet()->getCell("C$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="30%" align="center" > <?php echo $objPHPExcel->getActiveSheet()->getCell("D$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="15%" align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell("E$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="15%" align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell("F$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>

                                                <?php
                                                } else if ($i=='4' || $i=='6' || $i=='8' || $i=='10' || $i=='12' || $i=='14' || $i=='16' || $i=='18' || $i=='20' || $i=='22' ){
												?>
												<tr>
                                                <td align="center" width="5%" rowspan="2" > <b><?php echo $objPHPExcel->getActiveSheet()->getCell("B$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"');?></b></td>
                                                <td align="center" width="35%" rowspan="2" > <?php echo $objPHPExcel->getActiveSheet()->getCell("C$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="30%" align="center" > <?php echo $objPHPExcel->getActiveSheet()->getCell("D$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="15%" align="right" > <?php echo $objPHPExcel->getActiveSheet()->getCell("E$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="15%" align="right"> <?php echo $objPHPExcel->getActiveSheet()->getCell("F$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                  <?php
                                            		} else {
                                                ?>
                                                <tr>
                                               
                                                <td width="40%" align="center"> <?php echo $objPHPExcel->getActiveSheet()->getCell("D$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="15%" align="right"><?php echo $objPHPExcel->getActiveSheet()->getCell("E$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                <td width="15%" align="right" bgcolor="#2F353B"> <?php echo $objPHPExcel->getActiveSheet()->getCell("F$i")->getFormattedValue('#,##0,,;(#,##0,,);"-"'); ?></td>
                                                </tr>
                                                <?php
                                       			
                                       			}
                                               }
                                               ?>
                                                 
                                                </tbody>
                                            </table>
                                        </div>


                                               







                                    </div>
                                  
                                    
                                </div>
                            </div>
                            
                        </div>
                </div>

