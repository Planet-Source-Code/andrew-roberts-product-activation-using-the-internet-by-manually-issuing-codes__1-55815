<?PHP
    $fp = fopen("$productID//$licenseNo.txt","w");
    fputs($fp, "KTK,");
    fputs($fp, "$licenseNo,");
    fputs($fp, "$licenseHolder,");
    fputs($fp, "$licenseHardwareKey,");
    fputs($fp, "$daysEval,");
    fputs($fp, "$dateRegistered,");
    fputs($fp, "$lastUsed,");
    fputs($fp, "KTK");
    fclose($fp);

    echo "License Created.n";
    echo "$productID//$licenseNo.n";
    echo ".n";
    echo "If this license already existed then it was overwritten!";
?>