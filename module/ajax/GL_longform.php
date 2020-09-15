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
$tahun=$_POST['tahun'];
$bulan=$_POST['bulan'];

$tmp_tgl=$tahun."-$bulan-01";

//======================variable tanggal======================
$curr_tgl=date("Y-m-t",strtotime(date('Y-m-d',strtotime($tmp_tgl))." 0 second "));


$day=date('d',strtotime($curr_tgl));
$mon=date('M',strtotime($curr_tgl));
$year=date('y',strtotime($curr_tgl));

$label_tgl=$day."-".$mon."-".$year; // tanggal terpilih


$var_tabel=date('Ymd',strtotime($curr_tgl));
$table_asetkredit="DM_AsetKredit_$var_tabel";

$table_giro="DM_LiabilitasGiro_$var_tabel";
$table_tabungan="DM_LiabilitasTabungan_$var_tabel";
$table_deposito="DM_LiabilitasDeposito_$var_tabel";
$table_banklain="DM_LiabilitasKepadaBankLain_$var_tabel";

#############################################################################################


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
$objPHPExcel->getActiveSheet()->setCellValue('B2', "Export GL Longform $label_tgl ");

//GLOBAL

$objPHPExcel->getActiveSheet()->getStyle('A1:H1000')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('FFFFFF');

        $i=5;
        //$query_parameter=" select flash_level_3,flash_level_3_Description from Referensi_Flash_Report where flash_level_1_Description='$report_type' ";
        $query_parameter=" select Lap_Keu_Level_3,Lap_Keu_Level_3_Description from Referensi_Laporan_Keuangan where  ";
        $query_parameter.="  urut <> '99' order by urut asc ";
        //echo $query_parameter;
       // die();



        $result_parameter=odbc_exec($connection2, $query_parameter);
        while ( $row_param=odbc_fetch_array($result_parameter)) {

                $objPHPExcel->setActiveSheetIndex(0)->mergeCells("A$i:E$i");
                $objPHPExcel->getActiveSheet()->setCellValue("A$i", $row_param['Lap_Keu_Level_3_Description']);
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
                


                
                    switch ($row_param['Lap_Keu_Level_3']) {
                        
                        # Lap_Keu101000003  GIRO PADA BANK LAIN - PIHAK KETIGA
                        case 'Lap_Keu101000003':
                            $query_gl =" SELECT a.DataDate,a.kodegl,a.kodeproduct,a.kodecabang,SUM(Nominal) AS Nilai FROM DM_Journal a WITH (NOLOCK) ";
                            $query_gl.=" JOIN Referensi_GL_02 b ON b.GLNO = a.KodeGL AND b.PRODNO = a.KodeProduct ";
                            $query_gl.=" JOIN Referensi_Laporan_Keuangan c ON c.Lap_Keu_Level_3 = b.Lap_Keu_Level_3 ";
                            $query_gl.=" WHERE a.DataDate='$curr_tgl' AND b.Lap_Keu_Level_3='$row_param[Lap_Keu_Level_3]' and nominal > '0'  ";
                            $query_gl.=" group by a.DataDate,a.kodegl,a.kodeproduct,a.kodecabang ";
                            $result_gl=odbc_exec($connection2, $query_gl);
                            $found=odbc_num_rows($result_gl);
                            break;
                        
                        # Lap_Keu102000009  PINJAMAN YANG DITERIMA    
                        case 'Lap_Keu102000009':
                            $query_gl =" SELECT a.DataDate,a.kodegl,a.kodeproduct,a.kodecabang,SUM(Nominal) AS Nilai FROM DM_Journal a WITH (NOLOCK) ";
                            $query_gl.=" JOIN Referensi_GL_02 b ON b.GLNO = a.KodeGL AND b.PRODNO = a.KodeProduct ";
                            $query_gl.=" JOIN Referensi_Laporan_Keuangan c ON c.Lap_Keu_Level_3 = b.Lap_Keu_Level_3 ";
                            $query_gl.=" WHERE a.DataDate='$curr_tgl' AND b.Lap_Keu_Level_3='Lap_Keu101000003' and nominal < '0'  ";
                            $query_gl.=" group by a.DataDate,a.kodegl,a.kodeproduct,a.kodecabang ";
                            $result_gl=odbc_exec($connection2, $query_gl);
                            $found=odbc_num_rows($result_gl);
                            break;
                        
                        default:

                            $query_gl =" SELECT a.DataDate,a.kodegl,a.kodeproduct,a.kodecabang,SUM(Nominal) AS Nilai FROM DM_Journal a WITH (NOLOCK) ";
                            $query_gl.=" JOIN Referensi_GL_02 b ON b.GLNO = a.KodeGL AND b.PRODNO = a.KodeProduct ";
                            $query_gl.=" JOIN Referensi_Laporan_Keuangan c ON c.Lap_Keu_Level_3 = b.Lap_Keu_Level_3 ";
                            $query_gl.=" WHERE a.DataDate='$curr_tgl' AND b.Lap_Keu_Level_3='$row_param[Lap_Keu_Level_3]' ";
                            $query_gl.=" group by a.DataDate,a.kodegl,a.kodeproduct,a.kodecabang ";
                            $result_gl=odbc_exec($connection2, $query_gl);
                            $found=odbc_num_rows($result_gl);
                            break;
                    }

                
                    
                # x is start SUM
                $x=$i+1; 
                                while ( $row_gl=odbc_fetch_array($result_gl)) {
                                        $i++;
                                        $objPHPExcel->getActiveSheet()->setCellValue("A$i", $row_gl['DataDate']);
                                        $objPHPExcel->getActiveSheet()->setCellValue("B$i", $row_gl['kodegl']);
                                        $objPHPExcel->getActiveSheet()->setCellValue("C$i", $row_gl['kodeproduct']);
                                        $objPHPExcel->getActiveSheet()->setCellValue("D$i", $row_gl['kodecabang']);
                                        $objPHPExcel->getActiveSheet()->setCellValue("E$i", floatval($row_gl['Nilai']));
                        
                                }       # y is END SUM
                                        if ($found==0){
                                                $y=$i+1;
                                        }  else{
                                                $y=$i;
                                            }    
            
            
            $i++;
            #Black Fill
            $objPHPExcel->getActiveSheet()->getStyle("A$i:D$i")->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('000000');
            $objPHPExcel->getActiveSheet()->setCellValue("E$i", "=SUM(E$x:E$y)");

            $i++;    

        }

######## ADDITIONAL GL ##################
#- Aset Pihak Berelasi      
#- Aset Pihak Ketiga        
#- Liabilitas Pihak Berelasi       
#- Liabilitas Pihak Ketiga     

$array_add=array("1"=>"Aset Pihak Berelasi","2"=>"Aset Pihak Ketiga","3"=>"Liabilitas Pihak Berelasi","4"=>"Liabilitas Pihak Ketiga");

foreach ($array_add as $key => $value) {
  
                $objPHPExcel->setActiveSheetIndex(0)->mergeCells("A$i:E$i");
                $objPHPExcel->getActiveSheet()->setCellValue("A$i", $value);
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

                ### QUERY ##########

                switch ($key) {
                        
                        # Aset Pihak Berelasi
                        case '1':
                            $query_gl =" select DataDate,Managed_GL_Code,Managed_GL_Prod_Code,KodeCabang,sum(JumlahKreditPeriodeLaporan) as Nilai from $table_asetkredit ";
                            $query_gl.=" where DataDate='$curr_tgl' and Status_Pihak_Terkait like '%Y%' ";
                            $query_gl.=" group by DataDate,Managed_GL_Code,Managed_GL_Prod_Code,KodeCabang ";
                            $result_gl=odbc_exec($connection2, $query_gl);
                            $found=odbc_num_rows($result_gl);
                            break;
                        
                        # Aset Pihak Ketiga   
                        case '2':
                            $query_gl =" select DataDate,Managed_GL_Code,Managed_GL_Prod_Code,KodeCabang,sum(JumlahKreditPeriodeLaporan) as Nilai from $table_asetkredit ";
                            $query_gl.=" where DataDate='$curr_tgl' and Status_Pihak_Terkait like '%N%' ";
                            $query_gl.=" group by DataDate,Managed_GL_Code,Managed_GL_Prod_Code,KodeCabang ";
                            $result_gl=odbc_exec($connection2, $query_gl);
                            $found=odbc_num_rows($result_gl);
                            break;
                        
                         # Liabilitas Pihak Berelasi   
                        case '3':
                            $query_gl =" select DataDate,Managed_GL_Code,Managed_GL_Prod_Code,KodeCabang,sum(jumlahbulanlaporan) as Nilai from $table_giro ";
                            $query_gl.=" where DataDate='$curr_tgl'  and Status_Pihak_Terkait like '%Y%' group by DataDate,Managed_GL_Code,Managed_GL_Prod_Code,KodeCabang ";
                            $query_gl.=" union ";
                            $query_gl.=" select DataDate,Managed_GL_Code,Managed_GL_Prod_Code,KodeCabang,sum(jumlahbulanlaporan) as Nilai from $table_tabungan ";
                            $query_gl.=" where DataDate='$curr_tgl'  and Status_Pihak_Terkait like '%Y%' group by DataDate,Managed_GL_Code,Managed_GL_Prod_Code,KodeCabang ";
                            $query_gl.=" union ";
                            $query_gl.=" select DataDate,Managed_GL_Code,Managed_GL_Prod_Code,KodeCabang,sum(jumlahbulanlaporan) as Nilai from $table_deposito ";
                            $query_gl.=" where DataDate='$curr_tgl'  and Status_Pihak_Terkait like '%Y%' group by DataDate,Managed_GL_Code,Managed_GL_Prod_Code,KodeCabang ";
                            $query_gl.=" union ";
                            $query_gl.=" select DataDate,Managed_GL_Code,Managed_GL_Prod_Code,KodeCabang,sum(jumlahbulanlaporan) as Nilai from $table_banklain ";
                            $query_gl.=" where DataDate='$curr_tgl'  and Status_Pihak_Terkait like '%Y%' group by DataDate,Managed_GL_Code,Managed_GL_Prod_Code,KodeCabang ";
                            //echo $query_gl;
                            //die();
                            $result_gl=odbc_exec($connection2, $query_gl);
                            $found=odbc_num_rows($result_gl);
                            break;

                         # Liabilitas Pihak Ketiga   
                        case '4':
                            $query_gl =" select DataDate,Managed_GL_Code,Managed_GL_Prod_Code,KodeCabang,sum(jumlahbulanlaporan) as Nilai from $table_giro ";
                            $query_gl.=" where DataDate='$curr_tgl'  and Status_Pihak_Terkait like '%N%' group by DataDate,Managed_GL_Code,Managed_GL_Prod_Code,KodeCabang ";
                            $query_gl.=" union ";
                            $query_gl.=" select DataDate,Managed_GL_Code,Managed_GL_Prod_Code,KodeCabang,sum(jumlahbulanlaporan) as Nilai from $table_tabungan ";
                            $query_gl.=" where DataDate='$curr_tgl'  and Status_Pihak_Terkait like '%N%' group by DataDate,Managed_GL_Code,Managed_GL_Prod_Code,KodeCabang ";
                            $query_gl.=" union ";
                            $query_gl.=" select DataDate,Managed_GL_Code,Managed_GL_Prod_Code,KodeCabang,sum(jumlahbulanlaporan) as Nilai from $table_deposito ";
                            $query_gl.=" where DataDate='$curr_tgl'  and Status_Pihak_Terkait like '%N%' group by DataDate,Managed_GL_Code,Managed_GL_Prod_Code,KodeCabang ";
                            $query_gl.=" union ";
                            $query_gl.=" select DataDate,Managed_GL_Code,Managed_GL_Prod_Code,KodeCabang,sum(jumlahbulanlaporan) as Nilai from $table_banklain ";
                            $query_gl.=" where DataDate='$curr_tgl'  and Status_Pihak_Terkait like '%N%' group by DataDate,Managed_GL_Code,Managed_GL_Prod_Code,KodeCabang ";
                            //echo $query_gl;
                            //die();
                            $result_gl=odbc_exec($connection2, $query_gl);
                            $found=odbc_num_rows($result_gl);

                            break;
                       
                    }



                #### END QUERY #####        

                $x=$i+1; 
                                while ( $row_gl=odbc_fetch_array($result_gl)) {
                                        $i++;
                                        $objPHPExcel->getActiveSheet()->setCellValue("A$i", $row_gl['DataDate']);
                                        $objPHPExcel->getActiveSheet()->setCellValue("B$i", $row_gl['Managed_GL_Code']);
                                        $objPHPExcel->getActiveSheet()->setCellValue("C$i", $row_gl['Managed_GL_Prod_Code']);
                                        $objPHPExcel->getActiveSheet()->setCellValue("D$i", $row_gl['KodeCabang']);
                                        $objPHPExcel->getActiveSheet()->setCellValue("E$i", floatval($row_gl['Nilai']));
                        
                                }       # y is END SUM
                                        if ($found==0){
                                                $y=$i+1;
                                        }  else{
                                                $y=$i;
                                            }    
            
            
                $i++;
                 #Black Fill
                $objPHPExcel->getActiveSheet()->getStyle("A$i:D$i")->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('000000');
                $objPHPExcel->getActiveSheet()->setCellValue("E$i", "=SUM(E$x:E$y)");

$i++;

}



                


################################################

$objPHPExcel->getActiveSheet()->getStyle("A5:E$i")->applyFromArray($styleArrayBorder1);


// Rename sheet
$objPHPExcel->getActiveSheet()->setTitle("Export $report_type Longform");


$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
$objWriter->save("../report/download/GL_longform".$label_tgl."_".$file_eksport.".xls");


?>

                    <div class="portlet box blue" id="flash-report" >
                        <div class="portlet-title">
                            <div class="caption">
                                <i class="fa fa-book"></i> Result GL Longform
                            </div>
						
                        </div>
                        <div class="portlet-body">
                            <h4><b>PT Bank MNC Internasional, Tbk</b></h4>
                            <br>							
							<?php
							
							echo "<div class='alert alert-success'><strong> export GL '$report_type' Longform Success.... </div>";
							
							?>
                                <div class="tab-content">
                                    
                                   <div align="center" style="font-size:12px">
                                <a href="<?php echo "http://".$_SERVER['HTTP_HOST']."/fincon_dev/module/report/"."download/GL_longform".$label_tgl."_".$file_eksport.".xls";?>" 
                                    class="btn btn-sm green"> Download Excel  <i class="fa fa-arrow-circle-o-down"></i> </a> 
                            </div> 
                                    
                                </div>
                          
                            
                        </div>
                </div>

