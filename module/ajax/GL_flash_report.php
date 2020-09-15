<?php
//require_once 'config/config.php';
session_start();
require_once '../../config/config.php';
require_once '../../function/function.php';
require_once '../../session_login.php';
require_once '../../session_group.php';

require_once '../report/Classes/PHPExcel.php';
require_once '../report/Classes/PHPExcel/IOFactory.php';


date_default_timezone_set("Asia/Bangkok");
$file_eksport=date('Y_m_d_H_i_s');
logActivity("generate flash-report",date('Y_m_d_H_i_s'));

error_reporting(0);

//echo "oke........";
//die();
###################################

$tanggal=$_POST['tanggal']; 
$report_type=$_POST['report_type'];

$day=date('d',strtotime($tanggal));
$mon=date('M',strtotime($tanggal));
$year=date('y',strtotime($tanggal));

$label_tgl=$day."-".$mon."-".$year; // tanggal terpilih
$curr_tgl=date('Y-m-d',strtotime($tanggal));


        

// Create new PHPExcel object
$objPHPExcel = new PHPExcel();

$styleArrayFontBold = array('font' => array('bold'  => true,'color' => array('rgb' => ''),'size'  => 11,'name'  => 'Calibri'));
//$styleArraybackgroundRed = array('font' => array('bold'  => true,'color' => array('rgb' => ''),'size'  => 11,'name'  => 'Calibri'));
$styleArrayAlignmentCenter = array('alignment' => array(
            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
        ));
$styleArrayAlignmentCenter2 = array('alignment' => array(
            'vertical' => PHPExcel_Style_Alignment::VERTICAL_CENTER,
        ));
$styleArrayAlignmentRight= array('alignment' => array(
            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_RIGHT,
        ));
$styleArrayBorder1 = array('borders' => array('allborders' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('argb' => ''),),),);

//BOLD
$objPHPExcel->getActiveSheet()->getStyle('B1:B3')->applyFromArray($styleArrayFontBold);

//NUMBER FORMAT==================

$objPHPExcel->getActiveSheet()->getStyle('C8:H28')->getNumberFormat()->setFormatCode('#,##0_);(#,##0);"-"');



// 


//Bakgroud
//$objPHPExcel->getActiveSheet()->getStyle('B31:J33')->applyFromArray($styleArraybackgroundRed);
//$objPHPExcel->getActiveSheet()->getStyle('B5:J7')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('00FFFF');


//CENTER
$objPHPExcel->getActiveSheet()->getStyle('B5:J7')->applyFromArray($styleArrayAlignmentCenter2);

$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth(20);
// Create a first sheet, representing sales data
$objPHPExcel->setActiveSheetIndex(0);
$objPHPExcel->getActiveSheet()->setCellValue('B1', 'PT BANK MNC INTERNASIONAL TBK');
$objPHPExcel->getActiveSheet()->setCellValue('B2', "Export Flash Report $label_tgl ");

//GLOBAL

$objPHPExcel->getActiveSheet()->getStyle('A1:H1000')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');

        $i=5;
        $query_parameter=" select flash_level_3,flash_level_3_Description from Referensi_Flash_Report where flash_level_1_Description='$report_type' ";

        //echo $query_parameter;
       // die();
        $result_parameter=odbc_exec($connection2, $query_parameter);
        while ( $row_param=odbc_fetch_array($result_parameter)) {

                $objPHPExcel->setActiveSheetIndex(0)->mergeCells("A$i:E$i");
                $objPHPExcel->getActiveSheet()->setCellValue("A$i", $row_param['flash_level_3_Description']);
                $objPHPExcel->getActiveSheet()->getStyle("A$i:E$i")->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('808080');
                $objPHPExcel->getActiveSheet()->getStyle("A$i:E$i")->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
                $objPHPExcel->getActiveSheet()->getStyle("A$i:E$i")->applyFromArray($styleArrayFontBold);
                $i++;
                        $objPHPExcel->getActiveSheet()->setCellValue("A$i", "DataDate");
                        $objPHPExcel->getActiveSheet()->setCellValue("B$i", "KodeGL");
                        $objPHPExcel->getActiveSheet()->setCellValue("C$i", "KodeProduct");
                        $objPHPExcel->getActiveSheet()->setCellValue("D$i", "KodeCabang");
                        $objPHPExcel->getActiveSheet()->setCellValue("E$i", "Nominal");
                $objPHPExcel->getActiveSheet()->getStyle("A$i:E$i")->applyFromArray($styleArrayFontBold);
                //$objPHPExcel->getActiveSheet()->getStyle("A$i:E$i")->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('808080');


                
                    switch ($row_param['flash_level_3']) {
                        # FLASH101000007  Loans
                        case 'FLASH101000007':
                            $var_tabel=date('Ymd',strtotime($curr_tgl));
                            $table_asetkredit="DM_AsetKredit_$var_tabel";
                            $query_gl =" select DataDate,Managed_GL_Code,Managed_GL_Prod_Code,KodeCabang,sum (jumlahkreditperiodelaporan) as Nilai from $table_asetkredit ";
                            $query_gl.=" group by Managed_GL_Code,Managed_GL_Prod_Code,KodeCabang,DataDate ";
                            $result_gl=odbc_exec($connection2, $query_gl);
                            $found=odbc_num_rows($result_gl);
                            break;

                        # FLASH101000008  Performing Loan    
                        case 'FLASH101000008':
                            $var_tabel=date('Ymd',strtotime($curr_tgl));
                            $table_asetkredit="DM_AsetKredit_$var_tabel";
                            $query_gl =" select DataDate,Managed_GL_Code,Managed_GL_Prod_Code,KodeCabang,sum (jumlahkreditperiodelaporan) as Nilai  ";
                            $query_gl.=" from $table_asetkredit where kolektibilitas in ('1','2')  group by Managed_GL_Code,Managed_GL_Prod_Code,KodeCabang,DataDate ";
                            $result_gl=odbc_exec($connection2, $query_gl);
                            $found=odbc_num_rows($result_gl);
                            break;

                        # FLASH101000009  Non Performing Loan*)
                        case 'FLASH101000009':
                            $var_tabel=date('Ymd',strtotime($curr_tgl));
                            $table_asetkredit="DM_AsetKredit_$var_tabel";
                            $query_gl =" select DataDate,Managed_GL_Code,Managed_GL_Prod_Code,KodeCabang,sum (jumlahkreditperiodelaporan) as Nilai ";
                            $query_gl.=" from $table_asetkredit where kolektibilitas in ('3','4','5') group by Managed_GL_Code,Managed_GL_Prod_Code,KodeCabang,DataDate ";
                            $result_gl=odbc_exec($connection2, $query_gl);
                            $found=odbc_num_rows($result_gl);
                            break;

                        default:
                            $query_gl  =" SELECT a.DataDate,a.kodegl,a.kodeproduct,a.kodecabang,sum(a.nominal) AS Nilai FROM DM_Journal a WITH (NOLOCK) ";
                            $query_gl .=" left JOIN Referensi_GL_02 b ON b.GLNO = a.KodeGL AND b.PRODNO = a.KodeProduct ";
                            $query_gl .=" left JOIN Referensi_Flash_Report c ON c.FLASH_Level_3 = b.FLASH_LEVEL_3 ";
                            $query_gl .=" WHERE a.DataDate='$curr_tgl' and a. Flag_M='Y' and b.FLASH_Level_3='$row_param[flash_level_3]' ";
                            $query_gl .=" GROUP BY a.kodegl,a.kodeproduct,a.kodecabang,a.jenismatauang,a.DataDate ";
                            $query_gl .=" order by kodegl asc ";
                            $result_gl=odbc_exec($connection2, $query_gl);
                            $found=odbc_num_rows($result_gl);
                            break;
                    }

                /*
                    $query_gl  =" SELECT a.DataDate,a.kodegl,a.kodeproduct,a.kodecabang,sum(a.nominal) AS Nilai FROM DM_Journal a WITH (NOLOCK) ";
                    $query_gl .=" left JOIN Referensi_GL_02 b ON b.GLNO = a.KodeGL AND b.PRODNO = a.KodeProduct ";
                    $query_gl .=" left JOIN Referensi_Flash_Report c ON c.FLASH_Level_3 = b.FLASH_LEVEL_3 ";
                    $query_gl .=" WHERE a.DataDate='$curr_tgl' and a. Flag_M='Y' and b.FLASH_Level_3='$row_param[flash_level_3]' ";
                    $query_gl .=" GROUP BY a.kodegl,a.kodeproduct,a.kodecabang,a.jenismatauang,a.DataDate ";
                    $query_gl .=" order by kodegl asc ";
                */
                    

                 $x=$i+1; 
                    while ( $row_gl=odbc_fetch_array($result_gl)) {
                        $i++;
                        $objPHPExcel->getActiveSheet()->setCellValue("A$i", $row_gl['DataDate']);
                        $objPHPExcel->getActiveSheet()->setCellValue("B$i", $row_gl['kodegl']);
                        $objPHPExcel->getActiveSheet()->setCellValue("C$i", $row_gl['kodeproduct']);
                        $objPHPExcel->getActiveSheet()->setCellValue("D$i", $row_gl['kodecabang']);
                        $objPHPExcel->getActiveSheet()->setCellValue("E$i", floatval($row_gl['Nilai']));
                        
                        }
                            if ($found==0){
                                $y=$i+1;
                                 }else{
                                    $y=$i;
                                    }    
            
            # code...
            $i++;
            //$objPHPExcel->getActiveSheet()->setCellValue("E$i", "=SUM(E$x:$y)");
            $objPHPExcel->getActiveSheet()->getStyle("A$i:D$i")->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('000000');
            $objPHPExcel->getActiveSheet()->setCellValue("E$i", "=SUM(E$x:E$y)");

            $i++;    

        }

$objPHPExcel->getActiveSheet()->getStyle("A5:E$i")->applyFromArray($styleArrayBorder1);

//$objPHPExcel->setActiveSheetIndex(0)->mergeCells('B5:B7');//Account of Assets



// Rename sheet
$objPHPExcel->getActiveSheet()->setTitle("Export $report_type Flash");


$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
$objWriter->save("../report/download/GL_flash_report".$label_tgl."_".$file_eksport.".xls");


?>

                    <div class="portlet box blue" id="flash-report" >
                        <div class="portlet-title">
                            <div class="caption">
                                <i class="fa fa-book"></i> Result GL Flash Report
                            </div>
						
                        </div>
                        <div class="portlet-body">
                            <h4><b>PT Bank MNC Internasional, Tbk</b></h4>
                            <br>							
							<?php
							
							echo "<div class='alert alert-success'><strong> export GL '$report_type' Flash Report Success.... </div>";
							
							?>
                                <div class="tab-content">
                                    
                                   <div align="center" style="font-size:12px">
                                <a href="<?php echo "http://".$_SERVER['HTTP_HOST']."/fincon_dev/module/report/"."download/GL_flash_report".$label_tgl."_".$file_eksport.".xls";?>" 
                                    class="btn btn-sm green"> Download Excel  <i class="fa fa-arrow-circle-o-down"></i> </a> 
                            </div> 
                                    
                                </div>
                          
                            
                        </div>
                </div>

