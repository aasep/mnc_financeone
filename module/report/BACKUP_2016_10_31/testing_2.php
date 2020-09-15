<?php

function bd_nice_number($n) {
        // first strip any formatting;
        $n = (0+str_replace(",","",$n));
        
        // is this a number?
        if(!is_numeric($n)) return false;
        
        // now filter it;
        /*
        if($n>1000000000000) return round(($n/1000000000000),9).' trillion';
        else if($n>1000000000) return round(($n/1000000000),9).' billion';
        else if($n>1000000) return round(($n/1000000),9).' million';
        else if($n>1000) return round(($n/1000),9).' thousand';
        */
       // return number_format($n);
       return  number_format($n,2,",",".");
    }

echo bd_nice_number(23939393939.8745545);

//$price = '5.6078';
//$dec = 2;
//echo $price = number_format(floor($price*pow(10,$dec))/pow(10,$dec),$dec);

//$price = str_replace(',', '.', '0,10');
//echo number_format($price, 2, ',', '');



//echo(round(4.96754,3) . "<br>");
//echo(round(7.045,6) . "<br>");
//echo(round(7.055,2));


?>


