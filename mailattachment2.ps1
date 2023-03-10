$connectionString="Server=DLCTRLMNPDBAGL.jdadelivers.com;Database=ctm_automation;Integrated Security = false; User ID =hu_ctm_automation; Password = UL86IlDjSTIPxv4Jh8mo;"
#$Logfile ="C:\Users\1027315\Desktop\ControlM\error.csv"
$path = "C:\Users\1029452\Desktop\yamini\CONTROL-M\monthlymail.csv"
$date= Get-date -Format "dd/MM/yyyy"
$outputfile = "C:\Users\1029452\Desktop\yamini\CONTROL-M\monthlymail.xlsx"
#$inputfile ="C:\Users\1027315\Desktop\ControlM\Monthlydata.csv"

$lastMonth= (Get-Date).AddMonths(-1)
$today = $lastMonth
$Year = $today.Year
$Month = $today.Month



Function Excel_formatting{
Param (
        [Parameter(Mandatory=$true)]        
        [String]$outputfile,
        [Parameter(Mandatory=$true)]        
        [String] $inputfile
        
 )


$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false

$wb = $excel.Workbooks.Add()
$ws = $wb.Sheets.Item(1)

$ws.Cells.NumberFormat = "General"

write-output "Opening $inputfile"

$i = 1
Import-Csv $inputfile | Foreach-Object { 
    $j = 1
    foreach ($prop in $_.PSObject.Properties)
    {
        if ($i -eq 1) {
            $ws.Cells.Item($i, $j) = $prop.Name
            $head =$ws.cells.item($i, $j)
            $head.font.bold = $True
        } else {
            $ws.Cells.Item($i, $j) = $prop.Value
            $value =$ws.cells.item($i, $j)
            if($j -ne 1){
            if($prop.Value -eq 0 ){
            $value.Interior.ColorIndex  = 3}
            elseif($prop.Value -eq 1){
            $value.Interior.ColorIndex  = 10
            }
            }
        }
        $j++
    }
    $i++
}

$wb.SaveAs($outputfile)
$wb.Close()
$excel.Quit()
write-output "Success"
}


$JpowerID_group =(Get-details -connectionString $connectionString -sqlquery "select distinct(jpowerID) from SMReportingOutPut where Year = '$Year' and  Month = '$Month'").jpowerID 


 foreach ( $JpowerID in $JpowerID_group)
 { 

	# Insert data into Table1
Write-host "JpowerID $jpowerID"
#get The breach Information
$sla_BreachGP = (Get-Details -connectionString $connectionString  -sqlquery "select A.AccountName,SLA from SMReportingOutPut S join AccountDetails A on S.Year = $Year and S.Month = $Month and S.JpowerID =$JpowerID and S.JPowerID =A.JPowerID Order by Year,Month").SLA
$Customer = (Get-Details -connectionString $connectionString  -sqlquery "select A.AccountName,SLA from SMReportingOutPut S join AccountDetails A on S.$Year and S.$Month and S.JpowerID =$JpowerID and S.JPowerID =A.JPowerID Order by Year,Month").AccountName[0]
$Slapercentage = (Get-Details -connectionString $connectionString  -sqlquery "select Aggregate_SLA from MonthlyCompliance where  JpowerID =$JpowerID and Year = '$Year'and Month = '$Month'").Aggregate_SLA
  if($null -eq $Slapercentage){
   throw "SLA Monthly complainance value missing for the period '$Year' to '$Month'"
  }
      $hash=@{}   
      $hash = [pscustomobject]@{ AccountName = $Customer}
     $c=1
    foreach ($slavalue in $sla_BreachGP)
    {
   
    $hash | Add-Member -MemberType NoteProperty -Name "Day$c" -Value $slavalue
    $c++
    }
     $hash | Add-Member -MemberType NoteProperty -Name 'SLAPercentage' -Value "$Slapercentage%"
   
    $hash | Export-Csv -Path "$path" -Append -NoTypeInformation
	
	

}

Excel_formatting -outputfile $outputfile -inputfile $path
	$body = "Hi All,`n`n Please find the SLA batch compliance report for january-2021. So would request you to validate and confirm us.`n`n"
    $body += "Thank you,`n"
    $body += "Automation Team."
    $mail_subject = "Batch SLA Compliance Data ()"
	write-host "mail_subject: $mail_subject"   
	# Get mail parameters
	#$email_to= @('sunil.sharma@blueyonder.com','preetham.s@blueyonder.com')
    $email_to= @('yamini.gosula@blueyonder.com')
	$email_cc= @('yamini.gosula@blueyonder.com')
    $attachment = $Outputfile
	write-host "email_to : $email_to "
	write-host "email_cc : $email_cc"
	
			
			$MailParams = @{
				"From" = 'sla_management@jdadelivers.com'  
				"To" = $email_to
				"Cc"   = $email_cc
				"Subject" = "$mail_subject"
				"Body"    = $body
                "Attachment"= $attachment
				"SmtpServer" = 'mailout.jdadelivers.com'
			}
	write-host "MailParams : " $MailParams

		  
    #send mail
			
   try{
        Send-MailMessage  @MailParams #-BodyAsHtml
        return 1;
      }
   catch
     { $_.Exception;
      return $_.Exception
     }


     	
#$HtmlTable | ConvertTo-Csv -Path C:\Users\1027315\Desktop\ControlM\\FolderSizes.csv -NoTypeInformation

#$Path = "$env:temp\listOfServices.xlsx"
#Get-Service | Export-Excel -Path $Path -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -ClearSheet -WorksheetName 'List of Services' -Show
 # Export-Excel -Path $Path -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -ClearSheet -WorksheetName 'Report' -Show -Append


