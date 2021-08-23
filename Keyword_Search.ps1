#The strict mode limits cetain classes of bugs
Set-StrictMode -Version latest

#Prompt the user for their list of search terms. Could also be hard coded. 
$SearchTerms = Read-Host -Prompt "Provide search terms separated by a ','."

#Splits based on the ',' that is added from the user. I could do some input validation but seems like overkill at this point.
$TermArray = $SearchTerms.Split(",")

#Hard coded the path to eliminate any errors. Could be user inputed to provide flexibility.
#$path = "\\mpls-fs-100p\318iog\39 ios det 1"
$path = Read-Host -Prompt "Provide full path for directory to be searched."

#Script only works with MS Word and Excel files. PDFs require an external DLL(itextsharp.dll) from my research that would need to be downloded. 
$files = Get-Childitem -Path $path -Include *.docx,*.doc,*.xlsx,*.xlsm -Recurse -ErrorAction SilentlyContinue -Force | Where-Object { !($_.psiscontainer) }

#The below lines open never instances of the MS Word and Excel applications that will be used. 
$Word = New-Object -ComObject Word.Application
$Excel = New-Object -ComObject Excel.Application

#the visible value is set to false so as not to interfere with the user during execution.
$Word.visible = $False
$Excel.visible = $False

#Establish Word variables for searching
$matchCase = $false
$matchWholeWord = $true
$matchWildCards = $false
$matchSoundsLike = $false
$matchAllWordForms = $false
$forward = $true
$wrap = 1

#Need a dynamic storage variable and an ArrayList seemed to fit the need. A normal Array caused issues.
$results = New-Object -TypeName "System.Collections.ArrayList"

#Write-Output $results

#This is my first PS function and as such is not optimal, but is functional.
Function getStringMatch
{
    #The count is used to determine a progress status
    $status = $files.Count
    Write-Output "There are $status files to process."
    #Needed to count as the script progresses through the files.
    $i = 0

    # Loop through all files in the $path directory
    Foreach ($file In $files)
    {
        $i++
        Write-Progress -Activity "Processing files" -Status "Processing $($file)" -PercentComplete ($i/$files.Count * 100)

        #Write-Output $file
        #I used the If statements since there is a slight difference in how to handle MS Word vs Excel
        If($file -like "*.doc" -or $file -like "*.docx")
        {
            try{    
                $document = $Word.documents.open($file.FullName,$false,$true,$false,'ttt')
                $range = $document.content
            }catch{
            Continue
            }
            Foreach ($term In $TermArray)
                {
                    $wordFound = $range.find.execute($term,$matchCase,$matchWholeWord,$matchWildCards,
                                    $matchSoundsLike,$matchAllWordForms,$forward,$wrap)
                    If($wordFound)
                        {
                            #Hashtable to store the data for writing the csv output. Stored the data
             	            $properties = @{File_Path = $file.FullName; Matching_Term = $term}
                            $result = New-Object -TypeName PsCustomObject -Property $properties
		                    $results.Add($result)
                            break
	   	                }
	            }
            $document.close($false)

            }elseif($file -like "*.xlsx" -or $file -like "*.xlsm")
             {
                try{
                    $Workbook = $Excel.Workbooks.Open($file,$false,$true,5,'ttt')
                }catch{
                    Continue
                }
                ForEach($Sheet in $($Workbook.Sheets))
                    {
                        Foreach ($term in $TermArray)
                        {
                            try{
                            $Target = $Sheet.UsedRange.Find($term)
                            $try = $Target.Text
                            }catch{
                            Continue
                            }
                            if ($try -ilike $term)
                            {
                                $properties = @{File_Path = $file.FullName; Matching_Term = $term}
                                $result2 = New-Object -Typename PsCustomObject -Property $properties
                                $results.Add($result2)
                                break
                            }
                        }  
                    }
                $Workbook.close($false)	
            }

    }

    <#If($results)
    {
	    Foreach ($line in $results)
        {
		    $line | Add-Content -Path $output
	    }
    }#>

    if($results)
    {
        $results = $results | Sort-Object -Unique -Property File_Path
        $results | Export-Csv -Path $output -Delimiter ';' -NoTypeInformation -Encoding UTF8 -Force
    }else{
    write-host "Nothing was found! Have a good day"
    }

    $Word.quit()
    $Excel.quit()
}

#The file will currently save to the running users desktop.
$csvFileName = "PIIReturns.csv"
$output = "C:\Users\$env:USERNAME\Desktop\$csvFileName"

Set-Content $output -Value "Matching_Term,File_Path"
getStringMatch

