#The strict mode limits cetain classes of bugs
Set-StrictMode -Version latest

#Prompt the user for their list of search terms. Could also be hard coded. 
$SearchTerms = Read-Host -Prompt "Provide search terms separated by a ','."

#Splits based on the ',' that is added from the user. I could do some input validation but seems like overkill at this point.
$TermArray = $SearchTerms.Split(",")

#Hard coded the path to eliminate any errors. Could be user inputed to provide flexibility.

$path = Read-Host -Prompt "Provide full path for directory to be searched. Example:"

#Script only works with MS Word and Excel files. PDFs require an external DLL(itextsharp.dll) from my research that would need to be downloded. 
$files = Get-Childitem -Path $path -Include *.docx,*.doc,*.xlsx,*.xlsm -Recurse -ErrorAction SilentlyContinue -Force | Where-Object {!($_.psiscontainer)}
$wordFiles = New-Object -TypeName "System.Collections.ArrayList"
$excelFiles = New-Object -TypeName "System.Collections.ArrayList"

foreach ($file in $files){
    if ($file.Name -like "*.doc" -or $file -like "*.docx"){
        $wordFiles += $file
    }else{
        $excelFiles += $file
    }
}

#The below lines open never instances of the MS Word and Excel applications that will be used.
#Ran into issue trying to multithread excel. So I will process excel one at a time, but multithread word docs 
#$Word = New-Object -ComObject Word.Application
$Excel = New-Object -ComObject Excel.Application

#the visible value is set to false so as not to interfere with the user during execution.
#$Word.visible = $False
$Excel.visible = $False



#This is my first PS function and as such is not optimal, but is functional.

Function getStringMatch
{

    #Synchronized Hashtable 
    $config = [hashtable]::Synchronized(@{})
    $config.returns = New-Object -TypeName "System.Collections.ArrayList"


    # Define a script block to actually do the work
    $ScriptBlock = {
        Param($file, $TermArray, $config)

        #Establish Word variables for searching
        $matchCase = $false
        $matchWholeWord = $true
        $matchWildCards = $false
        $matchSoundsLike = $false
        $matchAllWordForms = $false
        $forward = $true
        $wrap = 1
        

        #The below lines open new instances of the MS Word and Excel applications that will be used. 
            $Word = New-Object -ComObject Word.Application
            $Word.visible = $False

            #the visible value is set to false so as not to interfere with the user during execution.
            
            try{
                #I needed to add the 'ttt' due some password protected files that stopped the script from running    
                $document = $Word.documents.open($file.FullName,$false,$true,$false,'ttt')
                $range = $document.content
            }catch{
                $Word.Quit()
            }
            Foreach ($term In $TermArray)
            {
                $wordFound = $range.find.execute($term,$matchCase,$matchWholeWord,$matchWildCards,
                                $matchSoundsLike,$matchAllWordForms,$forward,$wrap)
                If($wordFound)
                    {
                        #Hashtable to store the data for writing the csv output. Stored the data
             	        $properties = @{File = $file.FullName; Match = $term}
                        $result = New-Object -TypeName PsCustomObject -Property $properties
		                $config.returns += $result
                        Break
	   	            }
	        }
            $document.close($false)
            $Word.Quit()
    } #/ScriptBlock
    #The count is used to determine a progress status
    $status = $files.Count
    Write-Output "There are $status files to process."
    #Needed to count as the script progresses through the files.

    #Need a dynamic storage variable and an ArrayList seemed to fit the need. A normal Array caused issues.
    $results = New-Object -TypeName "System.Collections.ArrayList"

    $i = 0
    Foreach ($file in $excelFiles){
        $i++
        Write-Progress -Activity "Processing Excel files" -Status "Processing $($file)" -PercentComplete ($i/$excelFiles.Count * 100)
        try{
            #The ',' is added as an optional parameter for excel opening text files. 'ttt' is added to provide an automated input for password protected files.
            $Workbook = $Excel.Workbooks.Open($file,$false,$true,5,'ttt')
            }catch{
            Continue
        }
        ForEach ($Sheet in $($Workbook.Sheets))
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
                        $properties = @{File = $file.FullName; Match = $term}
                        $result2 = New-Object -Typename PsCustomObject -Property $properties
                        $results += $result2
                        Break
                    }
                }  
            }
        $Workbook.close($false)
    }
    $Excel.quit()

    # Create an empty arraylist that we'll use later
    $Jobs = New-Object -TypeName "System.Collections.ArrayList"

    # Create a Runspace Pool with a minimum and maximum number of run spaces. (http://msdn.microsoft.com/en-us/library/windows/desktop/dd324626(v=vs.85).aspx)
    $RunspacePool = [RunspaceFactory]::CreateRunspacePool(1,20)

    # Open the RunspacePool so we can use it
    $RunspacePool.Open()

    # Loop through all files in the $path directory
    $i = 0
    Foreach ($file In $wordFiles)
    {
        $i++
        Write-Progress -Activity "Processing Word files" -Status "Processing $($file)" -PercentComplete ($i/$wordFiles.Count * 100)

        $Powershell = [Powershell]::Create().AddScript($ScriptBlock)
        $Powershell.RunspacePool = $RunspacePool
        $Powershell.AddArgument($file).AddArgument($TermArray).addArgument($config) | Out-Null

        $JobObj = New-Object -TypeName PSObject -Property @{
            Runspace = $Powershell.BeginInvoke()
            Powershell = $Powershell
            } #/New-Object

        $Jobs.Add($JobObj) | Out-Null #added to a runspace arraylist
     } 
    #/ForEach}
        

    <#while ($Jobs.Runspace.IsCompleted -contains $false){
        Write-Host (Get-Date).ToString() "Still running..."
        Write-Host $Jobs.Count.ToString() " Jobs remaining"
        Start-Sleep 180
    }#>
    while($Jobs){
        Write-Host (Get-Date).ToString() "Still running..."
        Write-Host $Jobs.Count.ToString() " Jobs remaining"
        Start-Sleep 180
        foreach ($Runspace in $Jobs.ToArray()){
            If ($Runspace.Runspace.IsCompleted) {
                $Runspace.Powershell.Dispose()
                $Jobs.Remove($Runspace)
            }
        }
    }

    If($config.returns)
    {
        Foreach ($line in $config.returns)
        {
            $results += $line
        }
    }

    If($results)
    {
        $results = $results | Sort-Object -Unique -Property File
        Foreach ($line in $results)
        {
            $line | Add-Content -Path $output
        }
    }


}

#The file will currently save to the running users desktop.
$csvFileName = "PIIReturns.csv"
$output = "C:\Users\$env:USERNAME\Desktop\$csvFileName"

Set-Content $output -Value "Matching_Term,File_Path"
getStringMatch